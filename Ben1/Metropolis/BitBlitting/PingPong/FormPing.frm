VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   526
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKey 
      Height          =   495
      Left            =   6780
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox picPaddle 
      BackColor       =   &H0000FFFF&
      Enabled         =   0   'False
      Height          =   255
      Left            =   7200
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   2
      Top             =   1440
      Width           =   675
   End
   Begin VB.Timer tmrTime 
      Interval        =   20
      Left            =   240
      Top             =   120
   End
   Begin VB.PictureBox picBall 
      BackColor       =   &H00FF0000&
      Enabled         =   0   'False
      Height          =   675
      Left            =   7200
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   1
      Top             =   720
      Width           =   675
   End
   Begin VB.PictureBox picBlock 
      BackColor       =   &H000000FF&
      Enabled         =   0   'False
      Height          =   675
      Left            =   6600
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   0
      Width           =   1275
   End
   Begin VB.Line Line1 
      X1              =   436
      X2              =   436
      Y1              =   0
      Y2              =   384
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function BitBlt Lib "gdi32" _
                (ByVal hDestDC As Long, _
                 ByVal x As Long, _
                 ByVal Y As Long, _
                 ByVal nWidth As Long, _
                 ByVal nHeight As Long, _
                 ByVal hSrcDC As Long, _
                 ByVal xSrc As Long, _
                 ByVal ySrc As Long, _
                 ByVal dwRop As Long) As Long
Dim Playing As Boolean
Dim BlockX() As Integer, BlockY() As Integer, BlockBlit() As Boolean, Blocks As Integer, BlockPermanent() As Boolean
Dim BallX As Integer, BallY As Integer, XMove As Integer, YMove As Integer
Dim PaddleX As Integer, PaddleY As Integer, PaddleXMove As Integer, PaddleYMove As Integer
Dim PaddleLeft As Boolean, PaddleUp As Boolean, PaddleRight As Boolean, PaddleDown As Boolean
Private Sub Form_Load()
Dim i, x
Playing = True
tmrTime.Enabled = True
PaddleY = frmMain.ScaleHeight - 15
PaddleX = 300
PaddleXMove = 0
PaddleYMove = 0
BallX = 200
BallY = 100
XMove = 4
YMove = -4
For i = 0 To 164 Step 82
For x = 0 To 164 Step 41
Call GenerateBlock(i, x)
Next x
Next i
BlockPermanent(8) = True
BlockPermanent(10) = True
End Sub

Private Sub tmrTime_Timer()
Dim i
For i = 1 To Blocks
If BlockBlit(i) = True Then
If BallX > BlockX(i) - 41 And BallX < BlockX(i) + 82 Then
If BallY > BlockY(i) - 41 And BallY < BlockY(i) + 41 Then
If BallX < BlockX(i) Or BallX + 41 > BlockX(i) + 82 Then XMove = -XMove
YMove = -YMove
BallX = BallX + XMove
BallY = BallY + YMove
If BlockPermanent(i) = False Then BlockBlit(i) = False
End If
End If
End If
Next i
If BallX < 0 Or BallX > frmMain.ScaleWidth - 130 Then XMove = -XMove
If BallY < 0 Then YMove = -YMove
If BallX > PaddleX - 41 And BallX < PaddleX + 41 Then
If BallY > PaddleY - 41 And BallY < PaddleY + 41 Then
YMove = -YMove
End If
End If
BallX = BallX + XMove
BallY = BallY + YMove
If PaddleUp = True Then PaddleY = PaddleY - 5
If PaddleLeft = True Then PaddleX = PaddleX - 5
If PaddleDown = True Then PaddleY = PaddleY + 5
If PaddleRight = True Then PaddleX = PaddleX + 5
'PaddleX = PaddleX + PaddleXMove
'PaddleY = PaddleY + PaddleYMove
If PaddleX > frmMain.ScaleWidth - 130 Then PaddleX = frmMain.ScaleWidth - 130
If PaddleX < 0 Then PaddleX = 0
If PaddleY > frmMain.ScaleHeight - 15 Then PaddleY = frmMain.ScaleHeight - 15
If PaddleY < frmMain.ScaleHeight - 65 Then PaddleY = frmMain.ScaleHeight - 65
Me.Cls
For i = 1 To Blocks
If BlockBlit(i) = True Then BitBlt Me.hDC, BlockX(i), BlockY(i), picBlock.ScaleWidth, picBlock.ScaleHeight, picBlock.hDC, 0, 0, vbSrcCopy
Next i
BitBlt Me.hDC, BallX, BallY, picBall.ScaleWidth, picBall.ScaleHeight, picBall.hDC, 0, 0, vbSrcCopy
BitBlt Me.hDC, PaddleX, PaddleY, picPaddle.ScaleWidth, picPaddle.ScaleHeight, picPaddle.hDC, 0, 0, vbSrcCopy
Refresh
End Sub

Private Sub GenerateBlock(x, Y)
Blocks = Blocks + 1
ReDim Preserve BlockX(1 To Blocks)
ReDim Preserve BlockY(1 To Blocks)
ReDim Preserve BlockBlit(1 To Blocks)
ReDim Preserve BlockPermanent(1 To Blocks)
BlockX(Blocks) = x
BlockY(Blocks) = Y
BlockBlit(Blocks) = True
BlockPermanent(Blocks) = False
End Sub

Private Sub txtKey_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 37
PaddleLeft = True
Case Is = 38
PaddleUp = True
Case Is = 39
PaddleRight = True
Case Is = 40
PaddleDown = True
End Select

End Sub

Private Sub txtKey_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 37
PaddleLeft = False
Case Is = 38
PaddleUp = False
Case Is = 39
PaddleRight = False
Case Is = 40
PaddleDown = False
End Select
End Sub
