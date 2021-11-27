VERSION 5.00
Begin VB.Form frmBitBlt 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   Picture         =   "frmBitBlt.frx":0000
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMask 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2400
      ScaleHeight     =   78.546
      ScaleMode       =   0  'User
      ScaleWidth      =   70.054
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.Timer timerMove 
      Interval        =   100
      Left            =   1980
      Top             =   60
   End
   Begin VB.CommandButton cmdBitBlt 
      Caption         =   "Go!"
      Height          =   555
      Left            =   4020
      TabIndex        =   1
      Top             =   60
      Width           =   615
   End
   Begin VB.PictureBox picBitBlt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      ForeColor       =   &H000000FF&
      Height          =   1215
      Left            =   480
      Picture         =   "frmBitBlt.frx":BB844
      ScaleHeight     =   78.546
      ScaleMode       =   0  'User
      ScaleWidth      =   78.546
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmBitBlt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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

Private Sub cmdBitBlt_Click()
Me.Cls
BitBlt Me.hDC, 0, 0, picBitBlt.ScaleWidth, _
       picBitBlt.ScaleHeight, picBitBlt.hDC, 0, 0, vbSrcInvert
End Sub

Private Sub TimerMove_Timer()

Static X As Long, Y As Long


X = 40


'Keep the ball of the edge
If X > Me.ScaleWidth Then
    X = 0
End If

If Y > Me.ScaleHeight Then
    Y = 0
End If
Me.Cls
BitBlt Me.hDC, X, 90, picMask.ScaleWidth, picMask.ScaleHeight, picMask.hDC, 0, 0, vbSrcAnd
BitBlt Me.hDC, X, 90, picBitBlt.ScaleWidth, picBitBlt.ScaleHeight, picBitBlt.hDC, 0, 0, vbSrcPaint

End Sub


