VERSION 5.00
Begin VB.Form frmMaze 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   Caption         =   "Flazermaze"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   11535
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   4980
      Top             =   180
   End
   Begin VB.TextBox txtCommand 
      Height          =   375
      Left            =   -600
      TabIndex        =   0
      Top             =   660
      Width           =   195
   End
   Begin VB.Image imgLaser 
      Height          =   480
      Index           =   0
      Left            =   6780
      Picture         =   "frmMaze.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLPos 
      Height          =   480
      Index           =   7
      Left            =   4440
      Picture         =   "frmMaze.frx":08CA
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLPos 
      Height          =   480
      Index           =   6
      Left            =   3900
      Picture         =   "frmMaze.frx":1194
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLPos 
      Height          =   480
      Index           =   5
      Left            =   3360
      Picture         =   "frmMaze.frx":1A5E
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLPos 
      Height          =   480
      Index           =   4
      Left            =   2820
      Picture         =   "frmMaze.frx":2328
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLPos 
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "frmMaze.frx":2BF2
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLPos 
      Height          =   480
      Index           =   2
      Left            =   1740
      Picture         =   "frmMaze.frx":34BC
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLPos 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "frmMaze.frx":3D86
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLPos 
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "frmMaze.frx":4650
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   4
      Left            =   2820
      Picture         =   "frmMaze.frx":4F1A
      Top             =   1740
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "frmMaze.frx":57E4
      Top             =   1740
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   2
      Left            =   1740
      Picture         =   "frmMaze.frx":60AE
      Top             =   1740
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "frmMaze.frx":6978
      Top             =   1740
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEExplode 
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "frmMaze.frx":7242
      Top             =   1740
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFExplode 
      Height          =   480
      Index           =   4
      Left            =   2820
      Picture         =   "frmMaze.frx":7B0C
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFExplode 
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "frmMaze.frx":83D6
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFExplode 
      Height          =   480
      Index           =   2
      Left            =   1740
      Picture         =   "frmMaze.frx":8CA0
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFExplode 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "frmMaze.frx":956A
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFExplode 
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "frmMaze.frx":9E34
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEnemy 
      Height          =   480
      Index           =   0
      Left            =   7440
      Picture         =   "frmMaze.frx":A6FE
      Top             =   720
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEPos 
      Height          =   480
      Index           =   7
      Left            =   4440
      Picture         =   "frmMaze.frx":AFC8
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEPos 
      Height          =   480
      Index           =   6
      Left            =   3900
      Picture         =   "frmMaze.frx":B892
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEPos 
      Height          =   480
      Index           =   5
      Left            =   3360
      Picture         =   "frmMaze.frx":C15C
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEPos 
      Height          =   480
      Index           =   4
      Left            =   2820
      Picture         =   "frmMaze.frx":CA26
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEPos 
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "frmMaze.frx":D2F0
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEPos 
      Height          =   480
      Index           =   2
      Left            =   1740
      Picture         =   "frmMaze.frx":DBBA
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEPos 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "frmMaze.frx":E484
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgEPos 
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "frmMaze.frx":ED4E
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFPos 
      Height          =   480
      Index           =   7
      Left            =   4440
      Picture         =   "frmMaze.frx":F618
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFPos 
      Height          =   480
      Index           =   6
      Left            =   3900
      Picture         =   "frmMaze.frx":FEE2
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFPos 
      Height          =   480
      Index           =   5
      Left            =   3360
      Picture         =   "frmMaze.frx":107AC
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFPos 
      Height          =   480
      Index           =   4
      Left            =   2820
      Picture         =   "frmMaze.frx":11076
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFPos 
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "frmMaze.frx":11940
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFPos 
      Height          =   480
      Index           =   2
      Left            =   1740
      Picture         =   "frmMaze.frx":1220A
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFPos 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "frmMaze.frx":12AD4
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFPos 
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "frmMaze.frx":1339E
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape shpWall 
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgPlayer 
      Height          =   480
      Left            =   7440
      Picture         =   "frmMaze.frx":13C68
      Top             =   5520
      Width           =   480
   End
   Begin VB.Shape shpHalo 
      BorderColor     =   &H00800000&
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   5955
      Left            =   4740
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   5955
   End
End
Attribute VB_Name = "frmMaze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MoveFwd As Boolean, MoveBack As Boolean, SlideLeft As Boolean, SlideRight As Boolean, _
    TurnLeft As Boolean, TurnRight As Boolean, Running As Boolean
    Dim Walls As Integer
    Dim Enemies As Integer, EPos() As Integer, Vector() As Long, ExMove() As Integer, EyMove() As Integer
Dim Position As Integer, xMove As Integer, yMove As Integer
Dim EExplode(), PlayerExplode
Dim Lasers As Integer, Lasermin As Integer, LPos() As Integer
Dim LxMove() As Integer, LyMove() As Integer
Dim EGnrDamage(), ELsrDamage(), EFuseDamage()
Dim FGnrDamage, FLsrDamage, FFuseDamage
Private Sub Form_Load()
Dim i
Dim m, n, r, z, q
Call MakeWall(5940, 4860, 495, 1455)
Call MakeWall(8760, 4860, 495, 1875)
Call MakeWall(7080, 3600, 1035, 495)
Call MakeWall(7080, 7500, 975, 495)
Call MakeWall(5340, 3000, 495, 495)
Call MakeWall(9360, 3060, 4275, 495)
Call MakeWall(14700, 0, 495, 10995)
Call MakeWall(13140, 5040, 495, 2835)
Call MakeWall(10560, 5280, 1335, 495)
Call MakeWall(11220, 1080, 495, 2475)
Call MakeWall(0, 0, 15135, 495)
Call MakeWall(5820, 2520, 495, 495)
Call MakeWall(7260, 1800, 3075, 495)
'Call MakeWall(8580, 0, 495, 1815)
Call MakeWall(0, 10620, 15195, 495)
Call MakeWall(0, 60, 495, 11055)
Call MakeWall(3720, 3000, 495, 1935)
Call MakeWall(1800, 3660, 495, 2475)
Call MakeWall(1800, 7140, 4095, 495)
Call MakeWall(1800, 1920, 2415, 495)
'Call MakeWall(2700, 120, 495, 2295)
Call MakeWall(5460, 180, 495, 1395)
Call MakeWall(2880, 8280, 495, 495)
Call MakeWall(4140, 9360, 495, 495)
Call MakeWall(6540, 8880, 495, 495)
Call MakeWall(9780, 8580, 495, 2475)
Call MakeWall(10560, 6840, 1275, 795)
Call MakeWall(11640, 8880, 1995, 495)
Position = 0
Enemies = 2
Running = False
Lasermin = 1
Load imgEnemy(1)
With imgEnemy(1)
    .Top = 5460
    .Left = 1260
    .Visible = True
    .ZOrder
End With
Load imgEnemy(2)
With imgEnemy(2)
    .Top = 720
    .Left = 7440
    .Visible = True
    .ZOrder
End With
ReDim EPos(1 To Enemies)
ReDim ExMove(1 To Enemies)
ReDim EyMove(1 To Enemies)
ReDim Vector(1 To Enemies)
ReDim EExplode(1 To Enemies)
ReDim EGnrDamage(1 To Enemies)
ReDim ELsrDamage(1 To Enemies)
ReDim EFuseDamage(1 To Enemies)
For i = 1 To Enemies
EPos(i) = 0
Next i
imgPlayer.ZOrder
' Call MakeWall(0, 0, 1500, 495)
'  Call MakeWall(0, 0, 495, 1500)
' Call MakeWall(0, 500, 495, 495)
' Call MakeWall(1000, 500, 495, 495)
' Call MakeWall(2000, 500, 1500, 495)
End Sub

Private Sub tmrTime_Timer()
Dim i, x
'Lasers
If Lasers = 0 Then GoTo PlayerControls
For i = Lasermin To Lasers
If imgLaser(i).Visible = False Then
If i = Lasermin Then
Unload imgLaser(i)
Lasermin = Lasermin + 1
End If
GoTo Next_i
End If
Select Case LPos(i)
    Case Is = 0 'North
    LxMove(i) = 0
    LyMove(i) = 240
    Case Is = 1 'Northeast
    LxMove(i) = -180
    LyMove(i) = 45
    Case Is = 2 'East
    LxMove(i) = -240
    LyMove(i) = 0
    Case Is = 3 'Southeast
    LxMove(i) = -180
    LyMove(i) = -180
    Case Is = 4 'South
    LxMove(i) = 0
    LyMove(i) = -240
    Case Is = 5 'Southwest
    LxMove(i) = 180
    LyMove(i) = -180
    Case Is = 6 'West
    LxMove(i) = 240
    LyMove(i) = 0
    Case Is = 7 'Northwest
    LxMove(i) = 180
    LyMove(i) = 180
    End Select
    imgLaser(i).Move imgLaser(i).Left - LxMove(i), imgLaser(i).Top - LyMove(i)
    For x = 1 To Enemies
    If imgEnemy(x).Left > imgLaser(i).Left - 480 And _
imgEnemy(x).Left < imgLaser(i).Left + 480 Then
If imgEnemy(x).Top < imgLaser(i).Top + 480 And _
imgEnemy(x).Top > imgLaser(i).Top - 480 Then
EFuseDamage(x) = EFuseDamage(x) + 1
If EFuseDamage(x) = 20 Then EExplode(x) = True
imgLaser(i).Visible = False
End If
End If
Next x
Next_i:
Next i
PlayerControls:
'''
'''===<<<THE FOLLOWING CODE IS FOR THE PLAYER CONTROLS>>>===
'''
If PlayerExplode = True Then
Select Case imgPlayer.Picture
Case Is = imgFExplode(0).Picture
imgPlayer.Picture = imgFExplode(1).Picture
Case Is = imgFExplode(1).Picture
imgPlayer.Picture = imgFExplode(2).Picture
Case Is = imgFExplode(2).Picture
imgPlayer.Picture = imgFExplode(3).Picture
Case Is = imgFExplode(3).Picture
imgPlayer.Picture = imgFExplode(4).Picture
Case Else
If imgPlayer.Picture <> imgFExplode(4).Picture Then _
imgPlayer.Picture = imgFExplode(0).Picture
End Select
GoTo EnemyAI
End If
'''
'''
'''
If TurnLeft = True Then
If Position = 0 Then
Position = 7
Else
Position = Position - 1
End If
End If
If TurnRight = True Then
If Position = 7 Then
Position = 0
Else
Position = Position + 1
End If
End If
imgPlayer.Picture = imgFPos(Position).Picture
    If MoveFwd = True Then
    Select Case Position
    Case Is = 0 'North
    xMove = 0
    yMove = 60
    Case Is = 1 'Northeast
    xMove = -45
    yMove = 45
    Case Is = 2 'East
    xMove = -60
    yMove = 0
    Case Is = 3 'Southeast
    xMove = -45
    yMove = -45
    Case Is = 4 'South
    xMove = 0
    yMove = -60
    Case Is = 5 'Southwest
    xMove = 45
    yMove = -45
    Case Is = 6 'West
    xMove = 60
    yMove = 0
    Case Is = 7 'Northwest
    xMove = 45
    yMove = 45
    End Select
    Else
    xMove = 0
    yMove = 0
    End If
    If Running = True Then
    xMove = xMove * 2
    yMove = yMove * 2
    End If
For i = 1 To Walls
shpWall(i).Move shpWall(i).Left + xMove, shpWall(i).Top + yMove
Next i
For i = 1 To Enemies
imgEnemy(i).Move imgEnemy(i).Left + xMove, imgEnemy(i).Top + yMove
Next i
Call CheckCollision(Walls)
EnemyAI:
'''
'<<<===THIS IS THE CODE FOR THE ENEMY AI PLAYERS===>>>
'''
For i = 1 To Enemies
If EExplode(i) = True Then
Select Case imgEnemy(i).Picture
Case Is = imgEExplode(0).Picture
imgEnemy(i).Picture = imgEExplode(1).Picture
Case Is = imgEExplode(1).Picture
imgEnemy(i).Picture = imgEExplode(2).Picture
Case Is = imgEExplode(2).Picture
imgEnemy(i).Picture = imgEExplode(3).Picture
Case Is = imgEExplode(3).Picture
imgEnemy(i).Picture = imgEExplode(4).Picture
Case Else
If imgEnemy(i).Picture <> imgEExplode(4).Picture Then _
imgEnemy(i).Picture = imgEExplode(0).Picture
End Select
Exit Sub
End If
Call CheckEnemyCollision(Walls, i)
Select Case EPos(i)
    Case Is = 0 'North
    ExMove(i) = 0
    EyMove(i) = 60
    Case Is = 1 'Northeast
    ExMove(i) = -45
    EyMove(i) = 45
    Case Is = 2 'East
    ExMove(i) = -60
    EyMove(i) = 0
    Case Is = 3 'Southeast
    ExMove(i) = -45
    EyMove(i) = -45
    Case Is = 4 'South
    ExMove(i) = 0
    EyMove(i) = -60
    Case Is = 5 'Southwest
    ExMove(i) = 45
    EyMove(i) = -45
    Case Is = 6 'West
    ExMove(i) = 60
    EyMove(i) = 0
    Case Is = 7 'Northwest
    ExMove(i) = 45
    EyMove(i) = 45
    End Select
imgEnemy(i).Move imgEnemy(i).Left - ExMove(i), imgEnemy(i).Top - EyMove(i)
'
'vectors to friendly ships
'
If imgEnemy(i).Top = imgPlayer.Top Or imgEnemy(i).Left = imgPlayer.Left Then GoTo 1
Vector(i) = (imgEnemy(i).Top - imgPlayer.Top) / (imgEnemy(i).Left - imgPlayer.Left)
'For x = 1 To Walls
'm = shpWall(x).Height
'r = imgEnemy(i).Top - imgwall.Top
'q = Vector * shpWall(x).Left - imgEnemy(i).Left
'If q > x Or q < x - m Then
'Select Case Vector
'Case Is < 2 and > -2
1:
Next i
End Sub

Private Sub txtCommand_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 97 '1 = slide left
    SlideLeft = True
    Case Is = 98 '2 = move back
    MoveBack = True
    Case Is = 99 '3 = slide right
    SlideRight = True
    Case Is = 100 '4 = turn left
    TurnLeft = True
    Case Is = 102 '6 = turn right
    TurnRight = True
    Case Is = 104 '8 = move forward
    MoveFwd = True
    Case Is = 32 'SPACE = speed!
    Running = True
    Case Is = 13 'ENTER = laser
    Lasers = Lasers + 1
    Load imgLaser(Lasers)
    With imgLaser(Lasers)
    .Top = imgPlayer.Top
    .Left = imgPlayer.Left
    .Visible = True
    .Picture = imgLPos(Position)
    .ZOrder
    End With
    ReDim Preserve LPos(Lasermin To Lasers)
    ReDim Preserve LxMove(Lasermin To Lasers)
    ReDim Preserve LyMove(Lasermin To Lasers)
    LPos(Lasers) = Position
    End Select
   'MsgBox KeyCode
End Sub

Private Sub txtCommand_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case Is = 97 '1 = slide left
    SlideLeft = False
    Case Is = 98 '2 = move back
    MoveBack = False
    Case Is = 99 '3 = slide right
    SlideRight = False
    Case Is = 100 '4 = turn left
    TurnLeft = False
    Case Is = 102 '6 = turn right
    TurnRight = False
    Case Is = 104 '8 = move forward
    MoveFwd = False
    Case Is = 32 'SPACE = speed!
    Running = False
    End Select
End Sub

Private Sub MakeWall(x As Integer, Y As Integer, Width As Integer, Height As Integer)
Walls = Walls + 1
Load shpWall(Walls)
With shpWall(Walls)
    .Left = x
    .Top = Y
    .Height = Height
    .Width = Width
    .Visible = True
    .FillColor = frmMaze.BackColor
    .BorderColor = frmMaze.BackColor
    .ZOrder
End With
End Sub

Private Sub CheckCollision(WallNo As Integer)
Dim i
For i = 1 To WallNo
If imgPlayer.Left > shpWall(i).Left - 480 And _
imgPlayer.Left < shpWall(i).Left + shpWall(i).Width Then
If imgPlayer.Top < shpWall(i).Top + shpWall(i).Height And _
imgPlayer.Top > shpWall(i).Top - 480 Then
    Call collision
    Exit For
End If
End If
Next i
End Sub

Private Sub collision()
Dim i
For i = 1 To Walls
shpWall(i).Move shpWall(i).Left - xMove, shpWall(i).Top - yMove
Next i
For i = 1 To Enemies
imgEnemy(i).Move imgEnemy(i).Left - xMove, imgEnemy(i).Top - yMove
Next i
End Sub


Private Sub CheckEnemyCollision(WallNo, EnemyNo)
Dim i
For i = 1 To Walls
If imgEnemy(EnemyNo).Left > shpWall(i).Left - 480 And _
imgEnemy(EnemyNo).Left < shpWall(i).Left + shpWall(i).Width Then
If imgEnemy(EnemyNo).Top < shpWall(i).Top + shpWall(i).Height And _
imgEnemy(EnemyNo).Top > shpWall(i).Top - 480 Then
    Call EnemyCollision(EnemyNo)
    Exit For
End If
End If
Next i
If imgEnemy(EnemyNo).Left > imgPlayer.Left - 480 And _
imgEnemy(EnemyNo).Left < imgPlayer.Left + 480 Then
If imgEnemy(EnemyNo).Top < imgPlayer.Top + 480 And _
imgEnemy(EnemyNo).Top > imgPlayer.Top - 480 Then
EExplode(EnemyNo) = True
PlayerExplode = True
End If
End If
End Sub

Public Sub EnemyCollision(EnemyNo)
If EPos(EnemyNo) = 7 Then
EPos(EnemyNo) = 0
Else
EPos(EnemyNo) = EPos(EnemyNo) + 1
End If
imgEnemy(EnemyNo).Move imgEnemy(EnemyNo).Left + ExMove(EnemyNo), imgEnemy(EnemyNo).Top + EyMove(EnemyNo)
imgEnemy(EnemyNo).Picture = imgEPos(EPos(EnemyNo)).Picture
End Sub
