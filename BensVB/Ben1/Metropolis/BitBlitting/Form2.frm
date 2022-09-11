VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   LinkTopic       =   "Form2"
   ScaleHeight     =   304
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   697
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTime 
      Interval        =   50
      Left            =   5340
      Top             =   180
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   16
      Left            =   3840
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   16
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   15
      Left            =   3600
      Picture         =   "Form2.frx":0CCA
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   15
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   14
      Left            =   3360
      Picture         =   "Form2.frx":1994
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   14
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   13
      Left            =   3120
      Picture         =   "Form2.frx":265E
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   13
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   12
      Left            =   2880
      Picture         =   "Form2.frx":3328
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   12
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   11
      Left            =   2640
      Picture         =   "Form2.frx":3FF2
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   11
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   10
      Left            =   2400
      Picture         =   "Form2.frx":4CBC
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   10
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   9
      Left            =   2160
      Picture         =   "Form2.frx":5986
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   9
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   8
      Left            =   1920
      Picture         =   "Form2.frx":6650
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   8
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   7
      Left            =   1680
      Picture         =   "Form2.frx":731A
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   7
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   6
      Left            =   1440
      Picture         =   "Form2.frx":7FE4
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   6
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   5
      Left            =   1200
      Picture         =   "Form2.frx":8CAE
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   5
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   4
      Left            =   960
      Picture         =   "Form2.frx":9978
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   4
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   3
      Left            =   720
      Picture         =   "Form2.frx":A642
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   2
      Left            =   480
      Picture         =   "Form2.frx":B30C
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   2
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   1
      Left            =   240
      Picture         =   "Form2.frx":BFD6
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   1
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox picSmoke 
      Height          =   555
      Index           =   0
      Left            =   60
      Picture         =   "Form2.frx":CCA0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   60
      Width           =   555
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SmokeX(), SmokeY(), XAdd(), YAdd(), SmokeNo, FireX, FireY, FireNo()
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


Private Sub tmrTime_Timer()
Dim i
FireX = 0
FireY = 100
SmokeNo = SmokeNo + 1
ReDim FireNo(1 To 1)
If FireNo(1) = 15 Then
FireNo(1) = 0
Else
FireNo(1) = FireNo(1) + 1
End If
Me.Cls
BitBlt Me.hDC, FireX(i), FireY(i), picSmoke(0).ScaleWidth, picSmoke(0).ScaleHeight, picSmoke(FireNo(1)).hDC, 0, 0, vbSrcAnd

End Sub
