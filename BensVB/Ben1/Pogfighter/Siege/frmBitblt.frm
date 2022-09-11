VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11445
   LinkTopic       =   "Form2"
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   763
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTime 
      Interval        =   10
      Left            =   720
      Top             =   60
   End
End
Attribute VB_Name = "Form2"
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
Dim CursorX, CursorY
Dim Player As Player, Missile() As Missile, Missiles As Integer, MissileMin As Integer
Dim Enemy() As Enemy, Enemies As Integer, EnemyMin As Integer
Dim PlayerDC(1 To 1, 1 To 3, 1 To 10), MissileDC(1 To 1)
Dim EnemyDC(1 To 1, 1 To 3, 1 To 10), Kills As Integer

Private Sub Form_Load()
Dim i, j, k
Player.Mode = 1
Player.GraphicsDC = 1
Player.AniFrames(1) = 1
Player.MoveRate = 15
MissileMin = 1
EnemyMin = 1
For i = 1 To 1
    For j = 1 To 3
        For k = 1 To 10
            PlayerDC(i, j, k) = GenerateDC(App.Path & "\Player" & i & j & k & ".bmp")
            EnemyDC(i, j, k) = GenerateDC(App.Path & "\Enemy" & i & j & k & ".bmp")
        Next k
    Next j
Next i
For i = 1 To 1
MissileDC(i) = GenerateDC(App.Path & "\Missile" & i & ".bmp")
Next i
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Call LoadMissile(Player.x, Player.y, -6, 5, 0.1, 1)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
CursorX = x
CursorY = y
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub tmrTime_Timer()
Dim Distance As Double, Variation As Double, i, x
'<<<Generate Enemies>>>
If Int(Rnd * 15 + 1) = 15 Then
Call GenerateEnemy(1, 1, 1, 1, 6)
Enemy(Enemies).y = Int(Rnd * 500)
End If
'<<<Move Ship Towards Cursor>>>
Distance = Sqr((Player.y - CursorY) ^ 2 + (Player.x - CursorX) ^ 2)
If Distance < Player.MoveRate Then
Player.x = CursorX
Player.y = CursorY
End If
Variation = Distance / Player.MoveRate
If Variation = 0 Then Variation = 1
Player.x = Player.x - (Player.x - CursorX) / Variation
Player.y = Player.y - (Player.y - CursorY) / Variation
'<<<Update Ship Animation>>>
Player.CurFrame = Player.CurFrame + 1
If Player.CurFrame > Player.AniFrames(Player.Mode) Then Player.CurFrame = 1
'<<<Update Enemy Animation and Position>>>
If EnemyMin <= Enemies Then
For i = EnemyMin To Enemies
Enemy(i).x = Enemy(i).x + Enemy(i).Speed
Enemy(i).CurFrame = Enemy(i).CurFrame + 1
If Enemy(i).CurFrame > Enemy(i).AniFrames(Enemy(i).Mode) Then
If Enemy(i).Mode = 3 Then Enemy(i).x = 2001 'if explosion is complete, move enemy out of way
Enemy(i).CurFrame = 1
End If
'Delete old enemies
If i = EnemyMin And Enemy(i).x > 2000 Then EnemyMin = EnemyMin + 1
Next i
End If
'<<<Update Missile Position>>>
If MissileMin <= Missiles Then
For i = MissileMin To Missiles
Missile(i).YVar = Missile(i).YVar - Missile(i).YMinus
Missile(i).y = Missile(i).y - Missile(i).YVar
Missile(i).x = Missile(i).x + Missile(i).XVar
'<<<Check for Collisions with Enemies>>>
If EnemyMin <= Enemies Then
For x = EnemyMin To Enemies
If Missile(i).x > Enemy(x).x - 10 And Missile(i).x < Enemy(x).x + 60 And Missile(i).y > Enemy(x).y - 10 And Missile(i).y < Enemy(x).y + 60 Then
Missile(i).y = 1001
Enemy(x).Mode = 3 'explode enemy
Kills = Kills + 1
End If
Next x
End If
'Delete old missiles
If Missile(i).y > 1000 And i = MissileMin Then MissileMin = MissileMin + 1
Next i
End If
'<<<Blit To the Screen>>>
Me.Cls
BitBlt Me.hdc, Player.x - 30, Player.y - 30, 60, 60, PlayerDC(Player.GraphicsDC, Player.Mode, Player.CurFrame), 0, 0, vbSrcCopy
For i = MissileMin To Missiles
BitBlt Me.hdc, Missile(i).x - 5, Missile(i).y - 5, 10, 10, MissileDC(Missile(i).GraphicsDC), 0, 0, vbSrcCopy
Next i
For i = EnemyMin To Enemies
BitBlt Me.hdc, Enemy(i).x - 30, Enemy(i).y - 30, 60, 60, EnemyDC(Enemy(i).GraphicsDC, Enemy(i).Mode, Enemy(i).CurFrame), 0, 0, vbSrcCopy
Next i
End Sub

Private Sub LoadMissile(x, y, XVar, YVar, YMinus, GraphicsDC)
    Missiles = Missiles + 1
    ReDim Preserve Missile(1 To Missiles)
    Missile(Missiles).x = x
    Missile(Missiles).y = y
    Missile(Missiles).XVar = XVar
    Missile(Missiles).YVar = YVar
    Missile(Missiles).YMinus = YMinus
    Missile(Missiles).GraphicsDC = GraphicsDC
End Sub

Private Sub GenerateEnemy(GraphicsDC As Integer, Ani1, Ani2, Ani3, Speed)
Enemies = Enemies + 1
ReDim Preserve Enemy(1 To Enemies)
Enemy(Enemies).GraphicsDC = GraphicsDC
Enemy(Enemies).AniFrames(1) = Ani1
Enemy(Enemies).AniFrames(2) = Ani2
Enemy(Enemies).AniFrames(3) = Ani3
Enemy(Enemies).Speed = Speed
Enemy(Enemies).Mode = 1
End Sub
