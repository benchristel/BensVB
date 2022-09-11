VERSION 5.00
Begin VB.Form frmTrack1 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11115
   ClientLeft      =   15
   ClientTop       =   -585
   ClientWidth     =   15240
   ControlBox      =   0   'False
   Icon            =   "Speed_Demon_Course1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   12180
      Top             =   540
   End
   Begin VB.Timer tmrCountdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   540
      Top             =   120
   End
   Begin VB.Timer tmrMain 
      Interval        =   100
      Left            =   1020
      Top             =   120
   End
   Begin VB.TextBox txtMove 
      Height          =   435
      Left            =   8520
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   -600
      Width           =   150
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   16
      Left            =   2880
      Top             =   9600
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   15
      Left            =   13740
      Top             =   8940
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   14
      Left            =   12960
      Top             =   6780
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   13
      Left            =   6420
      Top             =   6780
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   12
      Left            =   4560
      Top             =   5040
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   11
      Left            =   5520
      Top             =   3900
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   9
      Left            =   2400
      Top             =   9000
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   8
      Left            =   5460
      Top             =   8040
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   7
      Left            =   6720
      Top             =   3780
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   6
      Left            =   9360
      Top             =   5160
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   5
      Left            =   10380
      Top             =   9720
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   4
      Left            =   13560
      Top             =   8520
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   3
      Left            =   13020
      Top             =   2160
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   2
      Left            =   12000
      Top             =   1320
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
      Left            =   12960
      TabIndex        =   2
      Top             =   480
      Width           =   1755
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
      TabIndex        =   1
      Top             =   1140
      Width           =   14655
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   10
      Left            =   1320
      Top             =   4020
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   0
      Left            =   1800
      Top             =   4440
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Shape shpWaypt 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   1
      Left            =   2580
      Top             =   1380
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgCar2 
      Height          =   480
      Index           =   1
      Left            =   840
      Picture         =   "Speed_Demon_Course1.frx":030A
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image imgCar2 
      Height          =   480
      Index           =   0
      Left            =   2580
      Picture         =   "Speed_Demon_Course1.frx":0614
      Top             =   3840
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   15
      Left            =   9900
      Picture         =   "Speed_Demon_Course1.frx":091E
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   14
      Left            =   9180
      Picture         =   "Speed_Demon_Course1.frx":0C28
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   13
      Left            =   8580
      Picture         =   "Speed_Demon_Course1.frx":0F32
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   12
      Left            =   7860
      Picture         =   "Speed_Demon_Course1.frx":123C
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   11
      Left            =   7200
      Picture         =   "Speed_Demon_Course1.frx":1546
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   10
      Left            =   6540
      Picture         =   "Speed_Demon_Course1.frx":1850
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   9
      Left            =   5880
      Picture         =   "Speed_Demon_Course1.frx":1B5A
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   8
      Left            =   5220
      Picture         =   "Speed_Demon_Course1.frx":1E64
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   7
      Left            =   4680
      Picture         =   "Speed_Demon_Course1.frx":216E
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   6
      Left            =   4020
      Picture         =   "Speed_Demon_Course1.frx":2478
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   5
      Left            =   3480
      Picture         =   "Speed_Demon_Course1.frx":2782
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   4
      Left            =   2880
      Picture         =   "Speed_Demon_Course1.frx":2A8C
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   3
      Left            =   2340
      Picture         =   "Speed_Demon_Course1.frx":2D96
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   2
      Left            =   1740
      Picture         =   "Speed_Demon_Course1.frx":30A0
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   1
      Left            =   1260
      Picture         =   "Speed_Demon_Course1.frx":33AA
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarPos 
      Height          =   480
      Index           =   0
      Left            =   720
      Picture         =   "Speed_Demon_Course1.frx":36B4
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCar1 
      Height          =   480
      Left            =   1740
      Picture         =   "Speed_Demon_Course1.frx":39BE
      Top             =   3240
      Width           =   480
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   10995
      Index           =   15
      Left            =   14700
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   14
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   10560
      Width           =   15135
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   5895
      Index           =   13
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   5295
      Index           =   12
      Left            =   3300
      Shape           =   4  'Rounded Rectangle
      Top             =   2580
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   5895
      Index           =   11
      Left            =   7740
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   5895
      Index           =   10
      Left            =   11340
      Shape           =   4  'Rounded Rectangle
      Top             =   2580
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   555
      Index           =   9
      Left            =   3300
      Shape           =   4  'Rounded Rectangle
      Top             =   7860
      Width           =   9675
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   8
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   10560
      Width           =   15135
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   5835
      Index           =   7
      Left            =   14700
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   5835
      Index           =   5
      Left            =   3300
      Shape           =   4  'Rounded Rectangle
      Top             =   2580
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   555
      Index           =   4
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2580
      Width           =   11715
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   15135
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   5160
      Width           =   8955
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   555
      Index           =   0
      Left            =   3780
      Shape           =   4  'Rounded Rectangle
      Top             =   2580
      Width           =   11415
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   14715
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   11055
      Index           =   6
      Left            =   60
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpIdentifier 
      BorderColor     =   &H0000FFFF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1740
      Shape           =   3  'Circle
      Top             =   3240
      Width           =   495
   End
   Begin VB.Shape shpFinish 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   540
      Top             =   4440
      Width           =   2775
   End
End
Attribute VB_Name = "frmTrack1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Accel, Brake, LeftMove, RightMove
Dim Speed, Pos
Dim Speed2(0 To 1), Pos2(0 To 1), NextWypt(0 To 1)
Dim Countdown, Lap, Place, Time, Time2(0 To 1), Lap2(0 To 1), Finished, CancelLap, CancelLap2(0 To 1)
Dim Score, Crashes
Dim Result
Private Sub Form_Load()
Dim i
frmTrack1.Visible = True
For i = 0 To 15
If shpRock(i).FillColor = &HE0E0E0 Then
shpRock(i).Left = shpRock(i).Left + 15330
shpRock(i).FillColor = &HC0C0C0
End If
Next i
For i = 2 To 10
shpWaypt(i).Left = shpWaypt(i).Left + 15330
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
NextWypt(i) = 1
Lap2(i) = 1
Speed2(i) = 0
Next i
Countdown = 6
Unload frmSelectTrack
Finished = False
End Sub

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
 End Select
lblCountdown.Caption = Countdown
End Sub

Private Sub tmrMain_Timer()
Dim i
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
If Finished = True Then Speed = 0
Speed = Speed - 3
Else
Speed = 0
End If
End If
If Speed2(0) < MaxSpeed + 5 Then Speed2(0) = Speed2(0) + AccelRate
If Speed2(1) < MaxSpeed Then Speed2(1) = Speed2(1) + AccelRate
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
    For i = 0 To 15
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
If imgCar1.Left > 13000 Then
For i = 0 To 15
shpRock(i).Left = shpRock(i).Left - Speed * 2
Next i
imgCar1.Left = imgCar1.Left - Speed * 2
For i = 0 To 16
shpWaypt(i).Left = shpWaypt(i).Left - Speed * 2
Next i
For i = 0 To 1
imgCar2(i).Left = imgCar2(i).Left - Speed * 2
Next i
shpFinish.Left = shpFinish.Left - Speed * 2
End If
If imgCar1.Left < 2330 Then
For i = 0 To 15
shpRock(i).Left = shpRock(i).Left + Speed * 2
Next i
imgCar1.Left = imgCar1.Left + Speed * 2
For i = 0 To 16
shpWaypt(i).Left = shpWaypt(i).Left + Speed * 2
Next i
For i = 0 To 1
imgCar2(i).Left = imgCar2(i).Left + Speed * 2
Next i
shpFinish.Left = shpFinish.Left + Speed * 2
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
If NextWypt(i) < 16 Then
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
Score = Time2(0) + Time2(1)
Score = Score - Time * 2
Select Case Place
Case Is = "1st"
Score = Int(Score * 1800)
Case Is = "2nd"
Score = Int(Score * 900)
Case Is = "3rd"
Score = 0
End Select
If Score < 0 Then Score = 0
MsgBox "Your score is " & Score & vbLf & "You placed " & Place & vbLf & "Your time was " & Int(Time) & " seconds" & vbLf & "You crashed " & Crashes & " time(s)" _
& vbLf & "Your average lap time was " & Int(Time / 3) & " seconds", 24, "Game Over!"
TotalScore = TotalScore + Score
Load frmSelectTrack
Unload frmTrack1
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
    Case Is = 27
    Result = MsgBox("Are you sure you want to abort this race?" & vbLf & "No score will be entered.", 48 + 4, "Exit Race")
    If Result = 6 Then Unload frmTrack1
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
