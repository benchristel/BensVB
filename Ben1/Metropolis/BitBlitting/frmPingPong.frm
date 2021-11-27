VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTime 
      Interval        =   50
      Left            =   60
      Top             =   60
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New Game"
      Height          =   435
      Left            =   4740
      TabIndex        =   2
      Top             =   0
      Width           =   675
   End
   Begin VB.PictureBox picMask 
      Height          =   660
      Left            =   4740
      Picture         =   "frmPingPong.frx":0000
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   1
      Top             =   480
      Width           =   660
   End
   Begin VB.PictureBox picBall 
      Height          =   600
      Left            =   4740
      Picture         =   "frmPingPong.frx":08CA
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   0
      Top             =   1200
      Width           =   660
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   312
      X2              =   312
      Y1              =   4
      Y2              =   216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BallX(1 To 4) As Integer, BallY(1 To 4) As Integer
Dim XMove(1 To 4) As Integer, YMove(1 To 4) As Integer
Dim MouseX, MouseY
Private Declare Function BitBlt Lib "gdi32" _
                (ByVal hDestDC As Long, _
                 ByVal X As Long, _
                 ByVal Y As Long, _
                 ByVal nWidth As Long, _
                 ByVal nHeight As Long, _
                 ByVal hSrcDC As Long, _
                 ByVal xSrc As Long, _
                 ByVal ySrc As Long, _
                 ByVal dwRop As Long) As Long

Private Sub cmdNewGame_Click()
Dim i
Me.Cls
For i = 1 To 4
XMove(i) = 3
YMove(i) = 3
BitBlt Me.hDC, BallX(i), BallY(i), picMask.ScaleWidth, picMask.ScaleHeight, picMask.hDC, 0, 0, vbSrcAnd
BitBlt Me.hDC, BallX(i), BallY(i), picBall.ScaleWidth, picBall.ScaleHeight, picBall.hDC, 0, 0, vbSrcPaint
Next i
Refresh
End Sub

Private Sub Form_Load()
BallX(1) = 0
BallY(1) = 0
BallX(2) = 0
BallY(2) = 50
BallX(3) = 0
BallY(3) = 100
BallX(4) = 200
BallY(4) = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseX = X
MouseY = Y
End Sub

Private Sub tmrTime_Timer()
Dim i, X
Me.Cls
For i = 1 To 4
BallX(i) = BallX(i) + XMove(i)
BallY(i) = BallY(i) + YMove(i)
If BallX(i) > 281 Or BallX(i) < 0 Then XMove(i) = -XMove(i)
If BallY(i) > 181 Or BallY(i) < 0 Then YMove(i) = -YMove(i)
If BallX(i) > MouseX - 32 And _
BallX(i) < MouseX + 32 Then
If BallY(i) < MouseX + 32 And _
BallY(i) > MouseX - 32 Then
XMove(i) = -XMove(i)
YMove(i) = -YMove(i)
End If
End If
BitBlt Me.hDC, BallX(i), BallY(i), picMask.ScaleWidth, picMask.ScaleHeight, picMask.hDC, 0, 0, vbSrcAnd
BitBlt Me.hDC, BallX(i), BallY(i), picBall.ScaleWidth, picBall.ScaleHeight, picBall.hDC, 0, 0, vbSrcPaint
BitBlt Me.hDC, MouseX, MouseY, picMask.ScaleWidth, picMask.ScaleHeight, picMask.hDC, 0, 0, vbSrcAnd
BitBlt Me.hDC, MouseX, MouseY, picBall.ScaleWidth, picBall.ScaleHeight, picBall.hDC, 0, 0, vbSrcPaint
For X = 1 To 4
If X = i Then GoTo 1
If BallX(i) > BallX(X) - 32 And _
BallX(i) < BallX(X) + 32 Then
If BallY(i) < BallY(X) + 32 And _
BallY(i) > BallY(X) - 32 Then
XMove(i) = -XMove(i)
YMove(i) = -YMove(i)
XMove(X) = -XMove(X)
YMove(X) = -YMove(X)
End If
End If
1:
Next X
Next i
Refresh
End Sub
