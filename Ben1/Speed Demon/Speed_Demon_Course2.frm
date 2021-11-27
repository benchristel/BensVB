VERSION 5.00
Begin VB.Form frmTrack2 
   BackColor       =   &H0000C0C0&
   Caption         =   "Speed Demons -- Desert Storm"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleMode       =   0  'User
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrDustStorm 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   120
   End
   Begin VB.TextBox txtMove 
      Height          =   435
      Left            =   8460
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   -600
      Width           =   150
   End
   Begin VB.Timer tmrMain 
      Interval        =   100
      Left            =   0
      Top             =   120
   End
   Begin VB.Timer tmrCountdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   12120
      Top             =   540
   End
   Begin VB.Image imgCar1 
      Height          =   480
      Left            =   4740
      Picture         =   "Speed_Demon_Course2.frx":0000
      Top             =   5640
      Width           =   480
   End
   Begin VB.Shape shpIdentifier 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   4740
      Shape           =   3  'Circle
      Top             =   5640
      Width           =   495
   End
   Begin VB.Shape shpDust 
      BorderColor     =   &H0000C0C0&
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   11055
      Left            =   -240
      Top             =   0
      Visible         =   0   'False
      Width           =   15195
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   23
      Left            =   4620
      Top             =   9900
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   22
      Left            =   7260
      Top             =   9240
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   21
      Left            =   8100
      Top             =   6180
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   20
      Left            =   13320
      Top             =   6480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   19
      Left            =   13020
      Top             =   -800
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   18
      Left            =   13020
      Top             =   780
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   17
      Left            =   8400
      Top             =   1380
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   16
      Left            =   8820
      Top             =   3480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblCountdown 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   5340
      Width           =   14655
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   555
      Index           =   21
      Left            =   10200
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   3615
      Index           =   20
      Left            =   12000
      Shape           =   4  'Rounded Rectangle
      Top             =   5340
      Width           =   555
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   3195
      Index           =   19
      Left            =   7260
      Shape           =   4  'Rounded Rectangle
      Top             =   7620
      Width           =   555
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   3495
      Index           =   18
      Left            =   2460
      Shape           =   4  'Rounded Rectangle
      Top             =   5340
      Width           =   555
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   555
      Index           =   17
      Left            =   2460
      Shape           =   4  'Rounded Rectangle
      Top             =   5340
      Width           =   10095
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   5175
      Index           =   16
      Left            =   7260
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   555
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   11295
      Index           =   6
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   -240
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   420
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   14715
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   555
      Index           =   0
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   2460
      Width           =   9075
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   6180
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   8955
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   3180
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   9195
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   555
      Index           =   4
      Left            =   9060
      Shape           =   4  'Rounded Rectangle
      Top             =   7020
      Width           =   3315
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   8175
      Index           =   5
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   2460
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   10635
      Index           =   7
      Left            =   14640
      Shape           =   4  'Rounded Rectangle
      Top             =   420
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   8
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   10560
      Width           =   9135
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   3615
      Index           =   9
      Left            =   6180
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   555
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   3735
      Index           =   10
      Left            =   9060
      Shape           =   4  'Rounded Rectangle
      Top             =   7020
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   3675
      Index           =   11
      Left            =   11880
      Shape           =   4  'Rounded Rectangle
      Top             =   7020
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   10695
      Index           =   12
      Left            =   14640
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   10755
      Index           =   13
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   180
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   14
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   10560
      Width           =   15135
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   555
      Index           =   15
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "Speed_Demon_Course2.frx":030A
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "Speed_Demon_Course2.frx":0614
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   2
      Left            =   1680
      Picture         =   "Speed_Demon_Course2.frx":091E
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   3
      Left            =   2280
      Picture         =   "Speed_Demon_Course2.frx":0C28
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   4
      Left            =   2820
      Picture         =   "Speed_Demon_Course2.frx":0F32
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   5
      Left            =   3420
      Picture         =   "Speed_Demon_Course2.frx":123C
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   6
      Left            =   3960
      Picture         =   "Speed_Demon_Course2.frx":1546
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   7
      Left            =   4620
      Picture         =   "Speed_Demon_Course2.frx":1850
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   8
      Left            =   5160
      Picture         =   "Speed_Demon_Course2.frx":1B5A
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   9
      Left            =   5820
      Picture         =   "Speed_Demon_Course2.frx":1E64
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   10
      Left            =   6480
      Picture         =   "Speed_Demon_Course2.frx":216E
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   11
      Left            =   7140
      Picture         =   "Speed_Demon_Course2.frx":2478
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   12
      Left            =   7800
      Picture         =   "Speed_Demon_Course2.frx":2782
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   13
      Left            =   8520
      Picture         =   "Speed_Demon_Course2.frx":2A8C
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   14
      Left            =   9120
      Picture         =   "Speed_Demon_Course2.frx":2D96
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   15
      Left            =   9840
      Picture         =   "Speed_Demon_Course2.frx":30A0
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCar2 
      Height          =   480
      Index           =   0
      Left            =   5340
      Picture         =   "Speed_Demon_Course2.frx":33AA
      Top             =   6300
      Width           =   480
   End
   Begin VB.Image imgCar2 
      Height          =   480
      Index           =   1
      Left            =   4080
      Picture         =   "Speed_Demon_Course2.frx":36B4
      Top             =   6300
      Width           =   480
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   1
      Left            =   12960
      Top             =   3540
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   0
      Left            =   4800
      Top             =   3960
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   2
      Left            =   13500
      Top             =   1560
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   3
      Left            =   2400
      Top             =   1140
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   4
      Left            =   1200
      Top             =   13000
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   5
      Left            =   2400
      Top             =   1800
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   6
      Left            =   5400
      Top             =   1860
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   7
      Left            =   6420
      Top             =   3600
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   8
      Left            =   1740
      Top             =   4140
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   9
      Left            =   1200
      Top             =   9300
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   10
      Left            =   5100
      Top             =   9840
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   11
      Left            =   6300
      Top             =   6360
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   12
      Left            =   9180
      Top             =   6720
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   13
      Left            =   11040
      Top             =   9600
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   14
      Left            =   13500
      Top             =   9780
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   15
      Left            =   13680
      Top             =   4200
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblLaps 
      BackStyle       =   0  'Transparent
      Caption         =   "Lap 1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   12900
      TabIndex        =   1
      Top             =   480
      Width           =   1755
   End
   Begin VB.Shape shpFinish 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   3600
      Top             =   8280
      Width           =   2775
   End
End
Attribute VB_Name = "frmTrack2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Accel, Brake, LeftMove, RightMove
Dim Speed, Pos
Dim Speed2(0 To 1), Pos2(0 To 1), NextWypt(0 To 1)
Dim Countdown, Lap, Place, Time, Time2(0 To 1), Lap2(0 To 1), Finished, CancelLap, CancelLap2(0 To 1)
Dim Score, Crashes, DustTime, DustStorm
Private Sub Form_Load()
Dim i
Randomize
frmTrack2.Visible = True
For i = 0 To 21
If shpRock(i).FillColor = &HE0E0E0 Then
shpRock(i).Visible = False
shpRock(i).FillColor = &HC0C0C0
End If
Next i
Brake = False
Accel = False
Speed = 0
tmrMain.Enabled = False
tmrCountdown.Enabled = True
Countdown = 6
Place = "1st"
Crashes = 0
Score = 0
Time = 0
Lap = 1
For i = 0 To 1
NextWypt(i) = 0
Lap2(i) = 1
Speed2(i) = 0
Next i
Countdown = 6
Unload frmSelectTrack
Finished = False
DustStorm = False
End Sub

'Private Sub mnuPause_Click()
'If tmrMain.Enabled = True Then
'tmrMain.Enabled = False
'mnuPause.Checked = True
'Exit Sub
'End If
'If tmrMain.Enabled = False Then
'tmrMain.Enabled = True
'mnuPause.Checked = False
'End If
'End Sub

Private Sub tmrCountdown_Timer()
Select Case Countdown
 Case Is = "Go!"
 Countdown = ""
 tmrCountdown.Enabled = False
 Case Is > 1
 Countdown = Countdown - 1
 Case Is = 1
 Countdown = "Go!"
 tmrMain.Enabled = True
 tmrDustStorm.Enabled = True
 DustTime = Int(Rnd * 40 + 20)
 End Select
lblCountdown.Caption = Countdown
End Sub

Private Sub tmrDustStorm_Timer()
DustTime = DustTime - 1
If DustTime = 0 And DustStorm = False Then
DustTime = 3
DustStorm = True
shpDust.Visible = True
End If
If DustTime = 0 And DustStorm = True Then
DustStorm = False
shpDust.Visible = False
DustTime = Int(Rnd * 40 + 20)
End If
End Sub

Private Sub tmrMain_Timer()
Dim i
If Finished = True Then Speed = 0
If Accel = True And Speed < MaxSpeed And Finished = False Then
Speed = Speed + AccelRate
End If
If LeftMove = True And Pos > 0 Then
Pos = Pos - 1
GoTo 1
End If
If LeftMove = True And Pos = 0 Then Pos = 15
1:
If RightMove = True And Pos = 15 Then
Pos = 0
GoTo 2
End If
If RightMove = True And Pos < 15 Then Pos = Pos + 1
2:
If Brake = True And Speed > 9 Then Speed = Speed - BrakeRate
If Accel = False Then
If Speed > 2 Then
Speed = Speed - 3
Else
Speed = 0
End If
End If
For i = 0 To 1
If Speed2(i) < MaxSpeed Then Speed2(i) = Speed2(i) + AccelRate
Next i
Select Case Pos
Case Is = 0
imgCar1.Top = imgCar1.Top - Speed * 2
imgCar1.Picture = imgCarPos(0).Picture
Case Is = 1
imgCar1.Top = imgCar1.Top - Speed * 1.5
imgCar1.Left = imgCar1.Left + Speed * 0.5
imgCar1.Picture = imgCarPos(1).Picture
Case Is = 2
imgCar1.Top = imgCar1.Top - Speed * 1
imgCar1.Left = imgCar1.Left + Speed * 1
imgCar1.Picture = imgCarPos(2).Picture
Case Is = 3
imgCar1.Top = imgCar1.Top - Speed * 0.5
imgCar1.Left = imgCar1.Left + Speed * 1.5
imgCar1.Picture = imgCarPos(3).Picture
Case Is = 4
imgCar1.Left = imgCar1.Left + Speed * 2
imgCar1.Picture = imgCarPos(4).Picture
Case Is = 5
imgCar1.Top = imgCar1.Top + Speed * 0.5
imgCar1.Left = imgCar1.Left + Speed * 1.5
imgCar1.Picture = imgCarPos(5).Picture
Case Is = 6
imgCar1.Top = imgCar1.Top + Speed * 1
imgCar1.Left = imgCar1.Left + Speed * 1
imgCar1.Picture = imgCarPos(6).Picture
Case Is = 7
imgCar1.Top = imgCar1.Top + Speed * 1.5
imgCar1.Left = imgCar1.Left + Speed * 0.5
imgCar1.Picture = imgCarPos(7).Picture
Case Is = 8
imgCar1.Top = imgCar1.Top + Speed * 2
imgCar1.Picture = imgCarPos(8).Picture
Case Is = 9
imgCar1.Top = imgCar1.Top + Speed * 1.5
imgCar1.Left = imgCar1.Left - Speed * 0.5
imgCar1.Picture = imgCarPos(9).Picture
Case Is = 10
imgCar1.Top = imgCar1.Top + Speed * 1
imgCar1.Left = imgCar1.Left - Speed * 1
imgCar1.Picture = imgCarPos(10).Picture
Case Is = 11
imgCar1.Top = imgCar1.Top + Speed * 0.5
imgCar1.Left = imgCar1.Left - Speed * 1.5
imgCar1.Picture = imgCarPos(11).Picture
Case Is = 12
imgCar1.Left = imgCar1.Left - Speed * 2
imgCar1.Picture = imgCarPos(12).Picture
Case Is = 13
imgCar1.Top = imgCar1.Top - Speed * 0.5
imgCar1.Left = imgCar1.Left - Speed * 1.5
imgCar1.Picture = imgCarPos(13).Picture
Case Is = 14
imgCar1.Top = imgCar1.Top - Speed * 1
imgCar1.Left = imgCar1.Left - Speed * 1
imgCar1.Picture = imgCarPos(14).Picture
Case Is = 15
imgCar1.Top = imgCar1.Top - Speed * 1.5
imgCar1.Left = imgCar1.Left - Speed * 0.5
imgCar1.Picture = imgCarPos(15).Picture
End Select
    For i = 0 To 21
        If shpRock(i).Visible = True Then
            If imgCar1.Top > shpRock(i).Top And imgCar1.Top < shpRock(i).Top + shpRock(i).Height Then
                If imgCar1.Left < shpRock(i).Left + shpRock(i).Width And imgCar1.Left > shpRock(i).Left Then
                imgCar1.Left = imgCar1.Left + Speed * 2
                imgCar1.Top = imgCar1.Top + Speed * 2
                Speed = 0
                Crashes = Crashes + 1
                End If
    '        If imgCar1.Left + 480 > shpRock(i).Left And imgCar1.Left < shpRock(i).Left Then
    '        imgCar1.Left = imgCar1.Left - Speed * 2
    '        imgCar1.Top = imgCar1.Top + Speed * 2
    '        Speed = 0
    '        End If
                If imgCar1.Left + 480 > shpRock(i).Left And imgCar1.Left + 480 < shpRock(i).Left + shpRock(i).Width Then
                imgCar1.Top = imgCar1.Top - Speed * 2
                imgCar1.Left = imgCar1.Left - Speed * 2
                Speed = 0
                Crashes = Crashes + 1
                End If
            End If
        End If
        If shpRock(i).Visible = True Then
            If imgCar1.Top + 480 > shpRock(i).Top And imgCar1.Top + 480 < shpRock(i).Top + shpRock(i).Height Then
                If imgCar1.Left < shpRock(i).Left + shpRock(i).Width And imgCar1.Left > shpRock(i).Left Then
                imgCar1.Left = imgCar1.Left + Speed * 2
                imgCar1.Top = imgCar1.Top - Speed * 2
                Crashes = Crashes + 1
                Speed = 0
                End If
    '        If imgCar1.Left + 480 > shpRock(i).Left And imgCar1.Left < shpRock(i).Left Then
    '        imgCar1.Left = imgCar1.Left - Speed * 2
    '        imgCar1.Top = imgCar1.Top + Speed * 2
    '        Speed = 0
    '        End If
                If imgCar1.Left + 480 > shpRock(i).Left And imgCar1.Left + 480 < shpRock(i).Left + shpRock(i).Width Then
                imgCar1.Top = imgCar1.Top - Speed * 2
                imgCar1.Left = imgCar1.Left - Speed * 2
                Speed = 0
                Crashes = Crashes + 1
                End If
            End If
            End If

    '        If imgCar1.Top < shpRock(i).Top + shpRock(i).Height And imgCar1.Top > shpRock(i).Top - 480 Then
    '            If imgCar1.Left < shpRock(i).Left + shpRock(i).Width And imgCar1.Left > shpRock(i).Left - 480 Then
    '                If imgCar1.Top >= shpRock(i).Top And imgCar1.Top < shpRock(i).Top + shpRock(i).Height Then
    '                imgCar1.Top = shpRock(i).Top + shpRock(i).Height + 150
    '                End If
    '                If imgCar1.Top + 480 > shpRock(i).Top And imgCar1.Top + 480 <= shpRock(i).Top + shpRock(i).Height Then
    '                shpRock(i).Top = imgCar1.Top + 480
    '                Speed = 0
    '                End If
    '                If imgCar1.Top < shpRock(i).Top And imgCar1.Top + 480 > shpRock(i).Top + shpRock(i).Width Then
    '                If imgCar1.Left + 480 >= shpRock(i).Left And imgCar1.Left + 480 < shpRock(i).Left + shpRock(i).Width Then
    '                imgCar1.Left = imgCar1.Left - 480
    '                Speed = 0
    '                Else
    '                imgCar1.Left = imgCar1.Left + 480
    '                Speed = 0
    '                End If
    '                End If
    '            End If
    '        End If


    Next i
If imgCar1.Top > 11520 Then
For i = 0 To 1
If imgCar2(i).Visible = False Then
imgCar2(i).Visible = True
Else
imgCar2(i).Visible = False
End If
Next i
If shpFinish.Visible = True Then
shpFinish.Visible = False
Else
shpFinish.Visible = True
End If
For i = 0 To 21
Select Case shpRock(i).Visible
Case Is = False
shpRock(i).Visible = True
Case Is = True
shpRock(i).Visible = False
End Select
Next i
imgCar1.Top = 0 + Speed
End If
If imgCar1.Top < 0 Then
For i = 0 To 1
If imgCar2(i).Visible = False Then
imgCar2(i).Visible = True
Else
imgCar2(i).Visible = False
End If
Next i
If shpFinish.Visible = True Then
shpFinish.Visible = False
Else
shpFinish.Visible = True
End If
For i = 0 To 21
Select Case shpRock(i).Visible
Case Is = False
shpRock(i).Visible = True
Case Is = True
shpRock(i).Visible = False
End Select
Next i
imgCar1.Top = 11520 - Speed
End If
For i = 0 To 1
If shpWaypt(NextWypt(i)).Left > imgCar2(i).Left Then
If shpWaypt(NextWypt(i)).Top < imgCar2(i).Top Then Pos2(i) = 2
If shpWaypt(NextWypt(i)).Top > imgCar2(i).Top Then Pos2(i) = 6
If shpWaypt(NextWypt(i)).Top < imgCar2(i).Top + 200 And shpWaypt(NextWypt(i)).Top > imgCar2(i).Top - 200 Then Pos2(i) = 4
End If
If shpWaypt(NextWypt(i)).Left < imgCar2(i).Left Then
If shpWaypt(NextWypt(i)).Top < imgCar2(i).Top Then Pos2(i) = 14
If shpWaypt(NextWypt(i)).Top > imgCar2(i).Top Then Pos2(i) = 10
If shpWaypt(NextWypt(i)).Top < imgCar2(i).Top + 200 And shpWaypt(NextWypt(i)).Top > imgCar2(i).Top - 200 Then Pos2(i) = 12
End If
If shpWaypt(NextWypt(i)).Left - 200 < imgCar2(i).Left And shpWaypt(NextWypt(i)).Left + 200 > imgCar2(i).Left Then
If shpWaypt(NextWypt(i)).Top < imgCar2(i).Top Then Pos2(i) = 0
If shpWaypt(NextWypt(i)).Top > imgCar2(i).Top Then Pos2(i) = 8
End If
If shpWaypt(NextWypt(i)).Left - 200 < imgCar2(i).Left And shpWaypt(NextWypt(i)).Left + 200 > imgCar2(i).Left And shpWaypt(NextWypt(i)).Top > imgCar2(i).Top - 200 And shpWaypt(NextWypt(i)).Top < imgCar2(i).Top + 200 Then
If NextWypt(i) < 23 Then
NextWypt(i) = NextWypt(i) + 1
Else
NextWypt(i) = 0
End If
End If
Select Case Pos2(i)
Case Is = 0
imgCar2(i).Top = imgCar2(i).Top - Speed2(i) * 2
imgCar2(i).Picture = imgCarPos(0).Picture
Case Is = 2
imgCar2(i).Top = imgCar2(i).Top - Speed2(i) * 1
imgCar2(i).Left = imgCar2(i).Left + Speed2(i) * 1
imgCar2(i).Picture = imgCarPos(2).Picture
Case Is = 4
imgCar2(i).Left = imgCar2(i).Left + Speed2(i) * 2
imgCar2(i).Picture = imgCarPos(4).Picture
Case Is = 6
imgCar2(i).Top = imgCar2(i).Top + Speed2(i) * 1
imgCar2(i).Left = imgCar2(i).Left + Speed2(i) * 1
imgCar2(i).Picture = imgCarPos(6).Picture
Case Is = 8
imgCar2(i).Top = imgCar2(i).Top + Speed2(i) * 2
imgCar2(i).Picture = imgCarPos(8).Picture
Case Is = 10
imgCar2(i).Top = imgCar2(i).Top + Speed2(i) * 1
imgCar2(i).Left = imgCar2(i).Left - Speed2(i) * 1
imgCar2(i).Picture = imgCarPos(10).Picture
Case Is = 12
imgCar2(i).Left = imgCar2(i).Left - Speed2(i) * 2
imgCar2(i).Picture = imgCarPos(12).Picture
Case Is = 14
imgCar2(i).Top = imgCar2(i).Top - Speed2(i) * 1
imgCar2(i).Left = imgCar2(i).Left - Speed2(i) * 1
imgCar2(i).Picture = imgCarPos(14).Picture
End Select
Next i
For i = 0 To 1
If imgCar2(i).Top > 11520 Then
imgCar2(i).Top = 0 + Speed
 NextWypt(i) = NextWypt(i) + 1
If imgCar2(i).Visible = True Then
imgCar2(i).Visible = False
Exit For
Else
imgCar2(i).Visible = True
Exit For
End If
End If
If imgCar2(i).Top < 0 Then
imgCar2(i).Top = 11520 - Speed
NextWypt(i) = NextWypt(i) + 1
If imgCar2(i).Visible = True Then
imgCar2(i).Visible = False
Exit For
Else
imgCar2(i).Visible = True
End If
End If
    If imgCar2(i).Top > shpFinish.Top And imgCar2(i).Top < shpFinish.Top + 495 And NextWypt(i) = 0 Then
        If imgCar2(i).Left > shpFinish.Left And imgCar2(i).Left < shpFinish.Left + 2775 Then
            If CancelLap2(i) = False Then
            Lap2(i) = Lap2(i) + 1
            End If
            CancelLap2(i) = True
            If Lap2(i) = 4 And Finished = False Then
            If Place = "1st" Then Place = "2nd"
            If Place = "2nd" Then Place = "3rd"
            End If
           
        End If
    Else
    CancelLap2(i) = False
    End If
Next i
shpIdentifier.Top = imgCar1.Top
shpIdentifier.Left = imgCar1.Left
    If imgCar1.Top > shpFinish.Top And imgCar1.Top < shpFinish.Top + 495 And shpFinish.Visible = True Then
        If imgCar1.Left > shpFinish.Left And imgCar1.Left < shpFinish.Left + 2775 Then
            If Pos = 13 Or Pos = 14 Or Pos = 15 Or Pos = 0 Or Pos = 1 Or Pos = 2 Or Pos = 3 Then
            If CancelLap = False Then
            Lap = Lap + 1
            CancelLap = True
            lblLaps.Caption = "Lap " & Lap
            End If
            If Lap = 4 Then
            lblCountdown.Caption = Place & " place!"
            Finished = True
            End If
            Else
            Pos = 0
            End If
        End If
    Else
    CancelLap = False
    End If
If Finished = True Then
If Lap2(0) >= 4 And Lap2(1) >= 4 Then
Score = (Time2(0) + Time2(1) - Time * 2) * MaxSpeed
Select Case Place
Case Is = "1st"
Score = Int(Score * 20)
Case Is = "2nd"
Score = Int(Score * 10)
Case Is = "3rd"
Score = 0
End Select
If Score < 0 Then Score = 0
MsgBox "Your score is " & Score & vbLf & "You placed " & Place & vbLf & "Your time was " & Int(Time) & " seconds" & vbLf & "You crashed " & Crashes & " time(s)" _
& vbLf & "Your average lap time was " & Int(Time / 3) & " seconds", 24, "Game Over!"
TotalScore = TotalScore + Score
Load frmSelectTrack
Unload frmTrack2
End If
End If
End Sub

Private Sub tmrTime_Timer()
Dim i
    If Lap < 4 Then Time = Time + 0.1
For i = 0 To 1
If Lap2(i) < 4 Then Time2(i) = Time2(i) + 0.1
Next i
End Sub

Private Sub txtMove_KeyDown(keycode As Integer, shift As Integer)
    Select Case keycode
    Case Is = 37
    LeftMove = True
    Case Is = 38
    Accel = True
    Case Is = 39
    RightMove = True
    Case Is = 40
    Brake = True
    End Select
End Sub
Private Sub txtmove_keyup(keycode As Integer, shift As Integer)
    Select Case keycode
    Case Is = 37
    LeftMove = False
    Case Is = 38
    Accel = False
    Case Is = 39
    RightMove = False
    Case Is = 40
    Brake = False
    End Select
End Sub


