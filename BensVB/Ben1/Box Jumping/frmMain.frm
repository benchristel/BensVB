VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   9300
      Top             =   1680
   End
   Begin VB.PictureBox picField 
      Height          =   7455
      Left            =   60
      ScaleHeight     =   534.707
      ScaleMode       =   0  'User
      ScaleWidth      =   600
      TabIndex        =   3
      Top             =   60
      Width           =   8115
   End
   Begin VB.Label lblLives 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lives: 3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8220
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   8220
      TabIndex        =   1
      Top             =   1140
      Width           =   2295
   End
   Begin VB.Label lblBoxes 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rockets: 0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   8220
      TabIndex        =   0
      Top             =   60
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
BoxDC(1) = GenerateDC(App.Path & "\Rocket1.bmp")
BoxDC(2) = GenerateDC(App.Path & "\Rocket2.bmp")
BoxMaskDC(1) = GenerateDC(App.Path & "\RocketMask1.bmp")
BoxMaskDC(2) = GenerateDC(App.Path & "\RocketMask2.bmp")
PlayerDC(1) = GenerateDC(App.Path & "\Duck1.bmp")
PlayerDC(2) = GenerateDC(App.Path & "\Duck2.bmp")
MaskDC(1) = GenerateDC(App.Path & "\DuckMask1.bmp")
MaskDC(2) = GenerateDC(App.Path & "\DuckMask2.bmp")
BackgroundDC = GenerateDC(App.Path & "\Background.bmp")
BackBuffDC = GenerateDC(App.Path & "\Background.bmp")
Lives = 3
BoxScore = 0
Score = 0
Boxes = 1
BoxMin = 1
ReDim Box(1 To 1)
Player.Jump = False
Player.OnBox = 0
Player.x = 300
Player.y = 250
Player.Position = 1
With Box(1)
    .DC = 1
    .Deleted = False
    .x = 250
    .y = 298
    .xMove = 0
    .yMove = 0
End With
End Sub

Private Sub Picture1_Click()

End Sub



Private Sub picField_KeyDown(KeyCode As Integer, Shift As Integer)
'37 = left
'39 = right
'38 = jump
Select Case KeyCode
Case Is = 37
Player.xMove = -5
Player.Position = 2
Case Is = 39
Player.xMove = 5
Player.Position = 1
Case Is = 38
If Player.Jump = False Then
Box(Player.OnBox).xMove = 0
Box(Player.OnBox).yMove = 2
Player.yMove = -15
Player.Jump = True
Player.OnBox = 0
Player.y = Player.y - 1
End If
End Select
End Sub

Private Sub picField_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 37
Player.xMove = Player.xMove = 0
Case Is = 39
Player.xMove = Player.xMove = 0
Case Is = 40
If Player.yMove > 0 Then Player.yMove = 0
End Select
End Sub

Private Sub tmrTime_Timer()
Dim i
Static timecount As Integer
timecount = timecount + 1
If timecount = 80 Then
timecount = 0
'generate boxes if 4 seconds have elapsed
Call GenerateBox
End If
If timecount = 40 Then
Boxes = Boxes + 1
ReDim Preserve Box(BoxMin To Boxes)
With Box(Boxes)
    .DC = 1
    .Deleted = False
    .y = Int(Rnd * 400)
    If Int(Rnd * 2) = 0 Then
    .xMove = -3
    .x = 600
    .DC = 2
    Else
    .xMove = 3
    .x = -130
    .DC = 1
    End If
End With
End If
'move boxes
For i = BoxMin To Boxes
If Box(i).Deleted = False Then
Box(i).x = Box(i).x + Box(i).xMove
Box(i).y = Box(i).y + Box(i).yMove
'score
If Box(i).y > picField.ScaleHeight Then
Score = Score + 50
BoxScore = BoxScore + 1
Box(i).Deleted = True
End If
End If
Next i
lblScore.Caption = Score
lblBoxes.Caption = "Rockets: " & BoxScore
'move player
If Player.OnBox = 0 Then
Player.x = Player.x + Player.xMove
Player.y = Player.y + Player.yMove
Else
Player.x = Player.x + Player.xMove + Box(Player.OnBox).xMove
Player.y = Player.y + Box(Player.OnBox).yMove
End If
If Player.Jump = True Then Player.yMove = Player.yMove + 1
'check for player collision with boxes
For i = BoxMin To Boxes
'if player intersects box then move player on top of box and
'make player's onBox set to index.
If Box(i).Deleted = False And Player.x <= Box(i).x + 96 And Player.x >= Box(i).x - 48 And Player.y <= Box(i).y And Player.y >= Box(i).y - 48 Then
Player.y = Box(i).y - 48
Player.yMove = 0
Player.Jump = False
Player.OnBox = i
GoTo esc 'ensure that player's onbox is not set to 0.
End If
Next i
Player.OnBox = 0
If Player.Jump = False Then
Player.Jump = True
Player.yMove = 2
End If
esc:
'blit player + boxes to backbuffer
BitBlt BackBuffDC, 0, 0, 600, 500, BackgroundDC, 0, 0, vbSrcCopy
For i = BoxMin To Boxes
If Box(i).Deleted = False Then BitBlt BackBuffDC, Box(i).x - 30, Box(i).y, 156, 48, BoxMaskDC(Box(i).DC), 0, 0, vbSrcAnd
If Box(i).Deleted = False Then BitBlt BackBuffDC, Box(i).x - 30, Box(i).y, 156, 48, BoxDC(Box(i).DC), 0, 0, vbSrcPaint
Next i
BitBlt BackBuffDC, Player.x, Player.y, 48, 48, MaskDC(Player.Position), 0, 0, vbSrcAnd
BitBlt BackBuffDC, Player.x, Player.y, 48, 48, PlayerDC(Player.Position), 0, 0, vbSrcPaint
'blit backbuffer to screen
BitBlt picField.hdc, 0, 0, 600, 500, BackBuffDC, 0, 0, vbSrcCopy
End Sub

