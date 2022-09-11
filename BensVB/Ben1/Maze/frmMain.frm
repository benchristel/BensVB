VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   FontTransparent =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6285
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picScreen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      DrawWidth       =   5
      Height          =   11055
      Index           =   2
      Left            =   7620
      ScaleHeight     =   733
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   2
      Top             =   0
      Width           =   7635
      Begin VB.Label lblAmmo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   435
         Index           =   2
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   435
         Index           =   2
         Left            =   6300
         TabIndex        =   3
         Top             =   0
         Width           =   1275
      End
   End
   Begin VB.PictureBox picScreen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C00000&
      DrawWidth       =   5
      Height          =   11055
      Index           =   1
      Left            =   0
      ScaleHeight     =   733
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   0
      Top             =   0
      Width           =   7635
      Begin VB.Timer tmrTime 
         Interval        =   50
         Left            =   2040
         Top             =   180
      End
      Begin VB.Label lblAmmo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   435
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblScore 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   435
         Index           =   1
         Left            =   6300
         TabIndex        =   1
         Top             =   0
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTSIZE As Long = &H40
Dim Wall() As Wall, Walls As Integer
Dim PlayerData(1 To 2) As Player, MoveRate As Integer
Dim ViewX(1 To 2), ViewY(1 To 2)
Dim PlayerDC(1 To 8), LaserDC(1 To 8), MaskDC(1 To 8), LMaskDC(1 To 8), AmmoDC, AMaskDC
Dim Bullet() As Bullet, Bullets As Integer, BulletMin As Integer
Dim AmmoBox() As AmmoBox
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 81 'Player1 sidestep left
PlayerData(1).Sidestep = -1
Case Is = 87 'Player1 move forward
PlayerData(1).Movement = 1
Case Is = 69 'player1 sidestep right
PlayerData(1).Sidestep = 1
Case Is = 65 'Player1 turn left
PlayerData(1).Position = PlayerData(1).Position - 1
If PlayerData(1).Position = 0 Then PlayerData(1).Position = 8
Case Is = 68 'Player1 turn right
PlayerData(1).Position = PlayerData(1).Position + 1
If PlayerData(1).Position = 9 Then PlayerData(1).Position = 1
Case Is = 88 'Player1 move backwards
PlayerData(1).Movement = -1
'Case Is = 83
'PlayerData(1).Running = 2
Case Is = 103 'Player2 sidestep left
PlayerData(2).Sidestep = -1
Case Is = 104 'player2 move forward
PlayerData(2).Movement = 1
Case Is = 105 'Player2 sidestep right
PlayerData(2).Sidestep = 1
Case Is = 100 'Player2 turn left
PlayerData(2).Position = PlayerData(2).Position - 1
If PlayerData(2).Position = 0 Then PlayerData(2).Position = 8
'Case Is = 101 'player2 run
'PlayerData(2).Running = 2
Case Is = 102 'player2 turn right
PlayerData(2).Position = PlayerData(2).Position + 1
If PlayerData(2).Position = 9 Then PlayerData(2).Position = 1
Case Is = 98 'player2 move backwards
PlayerData(2).Movement = -1
Case Is = 32 'Player1 fire
If PlayerData(1).Ammo > 0 Then
Call Shoot(PlayerData(1).x, PlayerData(1).y, PlayerData(1).Position)
PlayerData(1).Ammo = PlayerData(1).Ammo - 1
End If
Case Is = 13 'player2 fire
If PlayerData(2).Ammo > 0 Then
Call Shoot(PlayerData(2).x, PlayerData(2).y, PlayerData(2).Position)
PlayerData(2).Ammo = PlayerData(2).Ammo - 1
End If
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 81 'Player1 sidestep left
PlayerData(1).Sidestep = 0
Case Is = 87 'Player1 move forward
PlayerData(1).Movement = 0
Case Is = 69 'player1 sidestep right
PlayerData(1).Sidestep = 0
Case Is = 65 'Player1 turn left
PlayerData(1).Turning = 0
Case Is = 68 'Player1 turn right
PlayerData(1).Turning = 0
'Case Is = 83 'player1 run
'PlayerData(1).Running = 1
Case Is = 88 'Player1 move backwards
PlayerData(1).Movement = 0
Case Is = 103 'Player2 sidestep left
PlayerData(2).Sidestep = 0
Case Is = 104 'player2 move forward
PlayerData(2).Movement = 0
Case Is = 105 'Player2 sidestep right
PlayerData(2).Sidestep = 0
Case Is = 100 'Player2 turn left
PlayerData(2).Turning = 0
'Case Is = 101 'player2 run
'PlayerData(2).Running = 1
Case Is = 102 'player2 turn right
PlayerData(2).Turning = 0
Case Is = 98 'player2 move backwards
PlayerData(2).Movement = 0
End Select
End Sub

Private Sub Form_Load()
Dim i
Call InitializePlayers
Call BuildWalls
For i = 1 To 8
PlayerDC(i) = GenerateDC(App.Path & "\Player" & i & ".bmp")
LaserDC(i) = GenerateDC(App.Path & "\Laser" & i & ".bmp")
MaskDC(i) = GenerateDC(App.Path & "\Mask" & i & ".bmp")
LMaskDC(i) = GenerateDC(App.Path & "\LMask" & i & ".bmp")
AmmoDC = GenerateDC(App.Path & "\AmmoBox.bmp")
AMaskDC = GenerateDC(App.Path & "\AMask.bmp")
Next i
ReDim AmmoBox(1 To 4)
For i = 1 To 4
AmmoBox(i).State = 400
Next i
AmmoBox(1).x = 450
AmmoBox(1).y = 350
AmmoBox(2).x = 575
AmmoBox(2).y = 450
AmmoBox(3).x = 1125
AmmoBox(3).y = 450
AmmoBox(4).x = 600
AmmoBox(4).y = 975
MoveRate = 4
For i = 1 To 2
ViewX(i) = PlayerData(i).x - 240
ViewY(i) = PlayerData(i).y - 352
Next i
BulletMin = 1
End Sub
'IN: FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Private Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function


Private Sub InitializePlayers()
Dim i
For i = 1 To 2
PlayerData(i).Ammo = 12
PlayerData(i).Bombs = 0
PlayerData(i).Health = 0
PlayerData(i).Position = 1
PlayerData(i).Running = 1
Next i
PlayerData(1).x = 100
PlayerData(1).y = 100
PlayerData(2).x = 700
PlayerData(2).y = 700
End Sub

Private Sub BuildWall(x, y, Length, Position)
Walls = Walls + 1
ReDim Preserve Wall(1 To Walls)
Wall(Walls).Length = Length
Wall(Walls).Position = Position
Wall(Walls).x = x
Wall(Walls).y = y
End Sub

Private Sub TmrTime_Timer()
Dim i, x, BulletDelete As Integer, DeleteAdd As Boolean
Static XAdd(1 To 2), YAdd(1 To 2)
DeleteAdd = True
For i = BulletMin To Bullets
If Bullet(i).State = 1 Then
DeleteAdd = False
Select Case Bullet(i).Position
Case Is = 1
Bullet(i).y = Bullet(i).y - 32
Case Is = 2
Bullet(i).x = Bullet(i).x + 32
Bullet(i).y = Bullet(i).y - 32
Case Is = 3
Bullet(i).x = Bullet(i).x + 32
Case Is = 4
Bullet(i).x = Bullet(i).x + 32
Bullet(i).y = Bullet(i).y + 32
Case Is = 5
Bullet(i).y = Bullet(i).y + 32
Case Is = 6
Bullet(i).x = Bullet(i).x - 32
Bullet(i).y = Bullet(i).y + 32
Case Is = 7
Bullet(i).x = Bullet(i).x - 32
Case Is = 8
Bullet(i).x = Bullet(i).x - 32
Bullet(i).y = Bullet(i).y - 32
End Select
Else
If DeleteAdd = True Then BulletDelete = BulletDelete + 1
End If

For x = 1 To Walls 'check whether bullet has collided with a wall
If Wall(x).Position = 1 Then 'assuming the wall is vertical
    If Bullet(i).x > Wall(x).x - 31 And Bullet(i).x < Wall(x).x Then
        If Bullet(i).y > Wall(x).y - 31 And Bullet(i).y < Wall(x).y + Wall(x).Length Then
        Bullet(i).State = 0
        End If
    End If
Else ' Wall is horizontal
    If Bullet(i).x > Wall(x).x - 31 And Bullet(i).x < Wall(x).x + Wall(x).Length Then
        If Bullet(i).y > Wall(x).y - 31 And Bullet(i).y < Wall(x).y Then
        Bullet(i).State = 0
        End If
    End If
End If
Next x
For x = 1 To 2
    If PlayerData(x).x > Bullet(i).x - 31 And PlayerData(x).x < Bullet(i).x + 31 And Bullet(i).State = 1 Then
        If PlayerData(x).y > Bullet(i).y - 31 And PlayerData(x).y < Bullet(i).y + 31 Then
        Bullet(i).State = 0
        If x = 1 Then
        PlayerData(2).Health = PlayerData(2).Health + 1
        Else
        PlayerData(1).Health = PlayerData(1).Health + 1
        End If
        End If
    End If
Next x
Next i
BulletMin = BulletMin + BulletDelete
For i = 1 To 2
'Turn the player
'If PlayerData(i).Health > 0 Then
Select Case PlayerData(i).Position
Case Is = 1
XAdd(i) = 0
YAdd(i) = -2 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
XAdd(i) = XAdd(i) + 2 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
Case Is = 2
XAdd(i) = 1.4 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
YAdd(i) = -1.4 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
XAdd(i) = XAdd(i) + 1.4 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
YAdd(i) = YAdd(i) + 1.4 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
Case Is = 3
XAdd(i) = 2 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
YAdd(i) = 0
YAdd(i) = YAdd(i) + 2 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
Case Is = 4
XAdd(i) = 1.4 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
YAdd(i) = 1.4 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
XAdd(i) = XAdd(i) - 1.4 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
YAdd(i) = YAdd(i) + 1.4 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
Case Is = 5
XAdd(i) = 0
YAdd(i) = 2 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
XAdd(i) = XAdd(i) - 2 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
Case Is = 6
XAdd(i) = -1.4 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
YAdd(i) = 1.4 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
XAdd(i) = XAdd(i) - 1.4 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
YAdd(i) = YAdd(i) - 1.4 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
Case Is = 7
XAdd(i) = -2 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
YAdd(i) = 0
YAdd(i) = YAdd(i) - 2 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
Case Is = 8
XAdd(i) = -1.4 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
YAdd(i) = -1.4 * PlayerData(i).Movement * MoveRate * PlayerData(i).Running
XAdd(i) = XAdd(i) + 1.4 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
YAdd(i) = YAdd(i) - 1.4 * PlayerData(i).Sidestep * MoveRate * PlayerData(i).Running
End Select
For x = 1 To Walls
If Wall(x).Position = 1 Then 'assuming the wall is vertical
    If PlayerData(i).x + XAdd(i) > Wall(x).x - 31 And PlayerData(i).x + XAdd(i) < Wall(x).x Then
        If PlayerData(i).y + YAdd(i) > Wall(x).y - 31 And PlayerData(i).y + YAdd(i) < Wall(x).y + Wall(x).Length Then
        GoTo Collide
        End If
    End If
Else ' Wall is horizontal
    If PlayerData(i).x + XAdd(i) > Wall(x).x - 31 And PlayerData(i).x + XAdd(i) < Wall(x).x + Wall(x).Length Then
        If PlayerData(i).y + YAdd(i) > Wall(x).y - 31 And PlayerData(i).y + YAdd(i) < Wall(x).y Then
        GoTo Collide
        End If
    End If
End If
Next x
For x = 1 To 4
    If AmmoBox(x).x > PlayerData(i).x - 31 And AmmoBox(x).x < PlayerData(i).x + 31 And AmmoBox(x).State = 400 Then
        If AmmoBox(x).y > PlayerData(i).y - 31 And AmmoBox(x).y < PlayerData(i).y + 31 Then
        AmmoBox(x).State = 0
        PlayerData(i).Ammo = PlayerData(i).Ammo + 12
        End If
    End If
    If AmmoBox(x).State < 400 Then AmmoBox(x).State = AmmoBox(x).State + 1
Next x
ViewX(i) = ViewX(i) + XAdd(i)
PlayerData(i).x = PlayerData(i).x + XAdd(i)
ViewY(i) = ViewY(i) + YAdd(i)
PlayerData(i).y = PlayerData(i).y + YAdd(i)
'End If
Collide:
Next i
picScreen(1).Cls
picScreen(2).Cls
For i = 1 To 2
For x = 1 To Walls
If Wall(x).Position = 1 Then picScreen(i).Line (Wall(x).x - ViewX(i), Wall(x).y - ViewY(i))-(Wall(x).x - ViewX(i), Wall(x).y + Wall(x).Length - ViewY(i)), RGB(50, 0, 50)
If Wall(x).Position = 2 Then picScreen(i).Line (Wall(x).x - ViewX(i), Wall(x).y - ViewY(i))-(Wall(x).x + Wall(x).Length - ViewX(i), Wall(x).y - ViewY(i)), RGB(50, 0, 50)
Next x
For x = 1 To 2
BitBlt picScreen(i).hdc, PlayerData(x).x - ViewX(i), PlayerData(x).y - ViewY(i), 32, 32, MaskDC(PlayerData(x).Position), 0, 0, vbSrcAnd

BitBlt picScreen(i).hdc, PlayerData(x).x - ViewX(i), PlayerData(x).y - ViewY(i), 32, 32, PlayerDC(PlayerData(x).Position), 0, 0, vbSrcPaint
Next x
For x = BulletMin To Bullets
If Bullet(x).State = 1 Then
BitBlt picScreen(i).hdc, Bullet(x).x - ViewX(i), Bullet(x).y - ViewY(i), 32, 32, LMaskDC(Bullet(x).Position), 0, 0, vbSrcAnd
BitBlt picScreen(i).hdc, Bullet(x).x - ViewX(i), Bullet(x).y - ViewY(i), 32, 32, LaserDC(Bullet(x).Position), 0, 0, vbSrcPaint
End If
Next x
For x = 1 To 4
If AmmoBox(x).State = 400 Then
BitBlt picScreen(i).hdc, AmmoBox(x).x - ViewX(i), AmmoBox(x).y - ViewY(i), 32, 32, AMaskDC, 0, 0, vbSrcAnd
BitBlt picScreen(i).hdc, AmmoBox(x).x - ViewX(i), AmmoBox(x).y - ViewY(i), 32, 32, AmmoDC, 0, 0, vbSrcPaint
End If
Next x
Next i
'BitBlt picScreen(1).hdc, PlayerData(2).x - ViewX(1), PlayerData(2).y - ViewY(1), V
Refresh
For i = 1 To 2
lblScore(i).Caption = PlayerData(i).Health
lblAmmo(i).Caption = PlayerData(i).Ammo
Next i
End Sub

Private Sub txtKeys_KeyDown(KeyCode As Integer, Shift As Integer)
'8, W = move fwd
'2, X = move back
'4, A = turn left
'6, D = turn right
'7/9, Q/E = sidestep
'ENTER, SPACE = fire
MsgBox (KeyCode)
End Sub

Private Sub Shoot(x As Single, y As Single, Position As Integer)
Bullets = Bullets + 1
ReDim Preserve Bullet(1 To Bullets)
Bullet(Bullets).Position = Position
Select Case Position
Case Is = 1
Bullet(Bullets).x = x
Bullet(Bullets).y = y - 3
Case Is = 2
Bullet(Bullets).x = x + 3
Bullet(Bullets).y = y - 3
Case Is = 3
Bullet(Bullets).x = x + 3
Bullet(Bullets).y = y
Case Is = 4
Bullet(Bullets).x = x + 3
Bullet(Bullets).y = y + 3
Case Is = 5
Bullet(Bullets).x = x
Bullet(Bullets).y = y + 3
Case Is = 6
Bullet(Bullets).x = x - 3
Bullet(Bullets).y = y + 3
Case Is = 7
Bullet(Bullets).x = x - 3
Bullet(Bullets).y = y
Case Is = 8
Bullet(Bullets).x = x - 3
Bullet(Bullets).y = y - 3
End Select
Bullet(Bullets).State = 1 ' bullet is moving
End Sub

Private Sub BuildWalls()
Call BuildWall(0, 0, 5000, 1)
Call BuildWall(5000, 0, 5000, 1)
Call BuildWall(0, 0, 5000, 2)
Call BuildWall(0, 5000, 5000, 2)
Call BuildWall(200, 400, 300, 2)
Call BuildWall(500, 200, 600, 2)
Call BuildWall(500, 200, 200, 1)
Call BuildWall(800, 250, 250, 1)
Call BuildWall(550, 400, 100, 1)
Call BuildWall(550, 500, 250, 2)
Call BuildWall(1100, 200, 550, 1)
Call BuildWall(200, 400, 550, 1)
Call BuildWall(350, 500, 50, 2)
Call BuildWall(350, 850, 50, 2)
Call BuildWall(200, 950, 450, 2)
Call BuildWall(550, 600, 300, 1)
Call BuildWall(650, 950, 100, 1)
Call BuildWall(650, 1050, 150, 2)
Call BuildWall(800, 1050, 200, 1)
Call BuildWall(900, 1050, 200, 1)
Call BuildWall(900, 1050, 100, 2)
Call BuildWall(1000, 950, 100, 1)
Call BuildWall(1000, 950, 100, 2)
Call BuildWall(1100, 800, 150, 1)
Call BuildWall(800, 800, 300, 2)
Call BuildWall(800, 650, 150, 1)
Call BuildWall(750, 600, 350, 2)
Call BuildWall(600, 600, 100, 2)
Call BuildWall(650, 500, 100, 2)
'''<<<Map 2>>>'''

End Sub
