VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKeys 
      Height          =   495
      Left            =   4860
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   -1000
      Width           =   1215
   End
   Begin VB.PictureBox picField 
      Height          =   4095
      Left            =   60
      ScaleHeight     =   269
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   0
      Top             =   60
      Width           =   6015
      Begin VB.Timer tmrTime 
         Interval        =   100
         Left            =   5520
         Top             =   0
      End
   End
   Begin VB.Label lblKills 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   495
      Left            =   6120
      TabIndex        =   3
      Top             =   600
      Width           =   795
   End
   Begin VB.Label lblHP 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   60
      Width           =   795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Player As Player
Dim PlayerDC(1 To 6, 1 To 2) As Long, EnemyDC As Long
Dim MaskDC(1 To 6, 1 To 2) As Long
Dim BackBuffDC As Long, BackgroundDC As Long
Dim Enemy() As Enemy, enemies As Integer, enemymin As Integer
Dim Kills As Integer
Private Sub Form_Load()
Dim i, x
For i = 1 To 6
For x = 1 To 2
PlayerDC(i, x) = GenerateDC(App.Path & "\player" & i & x & ".bmp")
MaskDC(i, x) = GenerateDC(App.Path & "\mask" & i & x & ".bmp")
Next x
Next i
EnemyDC = GenerateDC(App.Path & "\enemy")
BackBuffDC = GenerateDC(App.Path & "\background.bmp")
BackgroundDC = GenerateDC(App.Path & "\background.bmp")
Player.Position = 1
Player.Mode = 1
Player.MaxHP = 20
Player.HP = 20
Player.x = 60
Player.y = 60
enemymin = 1
End Sub

Private Sub tmrTime_Timer()
Static runcycle
Dim i
If runcycle = 1 Then
runcycle = 2
Else
runcycle = 1
End If
'make player run
If Player.Mode = 3 Then
If Player.Position = 1 Then
Player.x = Player.x + 10
Else
Player.x = Player.x - 10
End If
End If
'generate enemies
If Int(Rnd * 30) = 0 Then
enemies = enemies + 1
ReDim Preserve Enemy(1 To enemies)
With Enemy(enemies)
    .HP = 12
    .Mode = 1
    .x = Int(Rnd * 200 + 20)
    .Position = Int(Rnd * 2 + 1)
End With
End If
'make enemies attack
For i = enemymin To enemies
If Abs(Enemy(i).x - Player.x) < 32 And Enemy(i).Mode <> 5 And Enemy(i).Mode <> 6 Then
Enemy(i).AttackDelay = Enemy(i).AttackDelay - 1
If Enemy(i).AttackDelay <= 0 Then
If Enemy(i).Mode = 4 Then
If Player.Mode <> 2 Then Player.HP = Player.HP - 1
Enemy(i).AttackDelay = 8
Enemy(i).Mode = 1
Else
Enemy(i).Mode = 4
Enemy(i).AttackDelay = 3
End If
End If
Else
If Enemy(i).Mode <> 5 And Enemy(i).Mode <> 6 Then Enemy(i).Mode = 1
End If
Next i
'blit operations
BitBlt BackBuffDC, 0, 0, 500, 400, BackgroundDC, 0, 0, vbSrcCopy
If runcycle = 1 Or Player.Mode <> 3 Then
BitBlt BackBuffDC, Player.x, 100, 32, 32, MaskDC(Player.Mode, Player.Position), 0, 0, vbSrcAnd
BitBlt BackBuffDC, Player.x, 100, 32, 32, PlayerDC(Player.Mode, Player.Position), 0, 0, vbSrcPaint
Else
BitBlt BackBuffDC, Player.x, 100, 32, 32, MaskDC(1, Player.Position), 0, 0, vbSrcAnd
BitBlt BackBuffDC, Player.x, 100, 32, 32, PlayerDC(1, Player.Position), 0, 0, vbSrcPaint
End If
For i = enemymin To enemies
If runcycle = 1 Or Enemy(i).Mode <> 3 Then
BitBlt BackBuffDC, Enemy(i).x, 100, 32, 32, MaskDC(Enemy(i).Mode, Enemy(i).Position), 0, 0, vbSrcAnd
BitBlt BackBuffDC, Enemy(i).x, 100, 32, 32, EnemyDC, 0, 0, vbSrcPaint
Else
BitBlt BackBuffDC, Enemy(i).x, 100, 32, 32, MaskDC(1, Enemy(i).Position), 0, 0, vbSrcAnd
BitBlt BackBuffDC, Enemy(i).x, 100, 32, 32, EnemyDC, 0, 0, vbSrcPaint
End If
If Enemy(i).Mode = 5 Then Enemy(i).Mode = 6
Next i
BitBlt picField.hdc, 0, 0, 500, 400, BackBuffDC, 0, 0, vbSrcCopy
txtKeys.SetFocus
lblHP.Caption = Player.HP
lblKills.Caption = Kills
End Sub

Private Sub txtKeys_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i
Select Case KeyCode
Case Is = 40 'down arrow
Player.Mode = 2
Case Is = 39 'rt arrow
If Player.Mode <> 2 And Player.Mode <> 4 Then Player.Mode = 3
Player.Position = 1
If Player.Mode = 4 Then
    For i = enemymin To enemies
        If Abs(Player.x - Enemy(i).x) < 32 Then Enemy(i).HP = Enemy(i).HP - 5
        If Enemy(i).HP <= 0 Then Enemy(i).Mode = 5 'dead
    Next i
End If
Case Is = 37 'left arrow
If Player.Mode <> 2 And Player.Mode <> 4 Then Player.Mode = 3
Player.Position = 2
Case Is = 13 ' enter
If Player.Mode <> 2 Then Player.Mode = 4
For i = enemymin To enemies
    If Abs(Player.x - Enemy(i).x) < 32 Then Enemy(i).HP = Enemy(i).HP - 5
    If Enemy(i).HP <= 0 And Enemy(i).Mode <> 6 Then
    Enemy(i).Mode = 5 'dead
    Kills = Kills + 1
    End If
Next i
End Select
'MsgBox KeyCode
End Sub


Private Sub txtKeys_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 40
Player.Mode = 1
Case Is = 39
If Player.Mode = 3 Then Player.Mode = 1
Case Is = 37
If Player.Mode = 3 Then Player.Mode = 1
Case Is = 13
If Player.Mode = 4 Then Player.Mode = 1
End Select
End Sub
