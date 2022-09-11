VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   14520
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   14520
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   120
      Top             =   480
   End
   Begin VB.TextBox txtKeys 
      Height          =   435
      Left            =   -1000
      TabIndex        =   0
      Top             =   600
      Width           =   315
   End
   Begin VB.Label lblThrottle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   10140
      Width           =   1215
   End
   Begin VB.Image imgRoll 
      Height          =   480
      Left            =   1320
      Picture         =   "frmFlightSim.frx":0000
      Top             =   10560
      Width           =   480
   End
   Begin VB.Image imgPitch 
      Height          =   480
      Left            =   1320
      Picture         =   "frmFlightSim.frx":08CA
      Top             =   10080
      Width           =   480
   End
   Begin VB.Label lblAltitude 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   60
      TabIndex        =   2
      Top             =   10620
      Width           =   1095
   End
   Begin VB.Label lblSpeed 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   10140
      Width           =   1095
   End
   Begin VB.Image imgPos 
      Height          =   480
      Index           =   7
      Left            =   3420
      Picture         =   "frmFlightSim.frx":1194
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPos 
      Height          =   480
      Index           =   6
      Left            =   2940
      Picture         =   "frmFlightSim.frx":1A5E
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPos 
      Height          =   480
      Index           =   5
      Left            =   2460
      Picture         =   "frmFlightSim.frx":2328
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPos 
      Height          =   480
      Index           =   4
      Left            =   1980
      Picture         =   "frmFlightSim.frx":2BF2
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPos 
      Height          =   480
      Index           =   3
      Left            =   1500
      Picture         =   "frmFlightSim.frx":34BC
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPos 
      Height          =   480
      Index           =   2
      Left            =   1020
      Picture         =   "frmFlightSim.frx":3D86
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPos 
      Height          =   480
      Index           =   1
      Left            =   540
      Picture         =   "frmFlightSim.frx":4650
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPos 
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "frmFlightSim.frx":4F1A
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgPlayer 
      Height          =   480
      Left            =   7260
      Picture         =   "frmFlightSim.frx":57E4
      Top             =   5160
      Width           =   480
   End
   Begin VB.Shape shpRunway 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   7155
      Left            =   7020
      Top             =   -1320
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Throttle, Elevator, Aileron, Rudder, Speed, Heading, Pitch, Roll, Altitude, Fuel
Dim ThrottleUp As Boolean, ThrottleDown As Boolean, _
AileronLeft As Boolean, AileronRight As Boolean, _
RudderLeft As Boolean, RudderRight As Boolean, _
ElevatorUp As Boolean, ElevatorDown As Boolean

Private Sub Form_Load()
Fuel = 1000000
Throttle = 0
End Sub

Private Sub tmrTime_Timer()
'<<<Throttle>>>
If ThrottleUp = True Then
Throttle = Throttle + 1
If Speed = 0 Then Speed = 1
End If
If ThrottleDown = True Then Throttle = Throttle - 1
If Throttle = -1 Then Throttle = 0
If Throttle = 101 Then Throttle = 100
Fuel = Fuel - Throttle
Speed = Speed + Throttle * 0.02 'accelerate or deccelerate by
' % of throttle
If Speed > 300 Then Speed = 300
If Speed < 75 And Altitude > 0 Then
If Pitch > -2 And Pitch < 3 Then
Pitch = Pitch - 1 'stalling speed
ElseIf Pitch <> -2 Then
Pitch = Pitch + 1
End If
End If
'<<<Ailerons>>>
If AileronLeft = True And Speed > 110 Then
If Roll > -3 Then
Roll = Roll - 1
Else
Roll = 4 'inverted
End If
End If
If AileronRight = True And Speed > 110 Then
If Roll < 4 Then
Roll = Roll + 1
Else
Roll = -3
End If
End If
'<<<Elevators>>>
If ElevatorUp = True And Speed > 95 Then
Select Case Roll
'''
Case Is = 0
If Pitch < 7 Then
Pitch = Pitch + 1
Else
Pitch = 0
End If
'''
Case Is = 1
If Pitch < 7 Then
Pitch = Pitch + 1
Else
Pitch = 0
End If
If Heading < 7 Then
Heading = Heading + 1
Else
Heading = 0
End If
Speed = Speed - 5
'''
Case Is = 2
If Heading < 7 Then
Heading = Heading + 1
Else
Heading = 0
End If
Speed = Speed - 2
'''
Case Is = 3
If Pitch > 0 Then
Pitch = Pitch - 1
Else
Pitch = 7
End If
If Heading < 7 Then
Heading = Heading + 1
Else
Heading = 0
End If
Speed = Speed - 1
'''
Case Is = 4 'inverted
If Pitch > 0 Then
Pitch = Pitch - 1
Else
Pitch = 7
End If
'''
Case Is = -3
If Pitch > 0 Then
Pitch = Pitch - 1
Else
Pitch = 7
End If
If Heading > 0 Then
Heading = Heading - 1
Else
Heading = 7
End If
Speed = Speed - 1
'''
Case Is = -2
If Heading > 0 Then
Heading = Heading - 1
Else
Heading = 7
End If
'''
Case Is = -1
If Pitch > 0 Then
Pitch = Pitch - 1
Else
Pitch = 0
End If
If Heading > 0 Then
Heading = Heading - 1
Else
Heading = 7
End If
Speed = Speed - 2
End Select
End If
If ElevatorDown = True And Speed > 95 And Altitude > 0 Then
Select Case Roll
'''
Case Is = 0
If Pitch > 0 Then
Pitch = Pitch - 1
Else
Pitch = 7
End If
'''
Case Is = 1
If Pitch > 0 Then
Pitch = Pitch - 1
Else
Pitch = 7
End If
If Heading > 0 Then
Heading = Heading - 1
Else
Heading = 7
End If
Speed = Speed - 5
'''
Case Is = 2
If Heading > 0 Then
Heading = Heading - 1
Else
Heading = 7
End If
Speed = Speed - 2
'''
Case Is = 3
If Pitch < 7 Then
Pitch = Pitch + 1
Else
Pitch = 0
End If
If Heading > 0 Then
Heading = Heading - 1
Else
Heading = 7
End If
Speed = Speed - 1
'''
Case Is = 4 'inverted
If Pitch < 7 Then
Pitch = Pitch + 1
Else
Pitch = 0
End If
'''
Case Is = -3
If Pitch < 7 Then
Pitch = Pitch + 1
Else
Pitch = 0
End If
If Heading > 7 Then
Heading = Heading + 1
Else
Heading = 0
End If
Speed = Speed - 1
'''
Case Is = -2
If Heading < 7 Then
Heading = Heading + 1
Else
Heading = 0
End If
'''
Case Is = -1
If Pitch > 0 Then
Pitch = Pitch + 1
Else
Pitch = 7
End If
If Heading < 7 Then
Heading = Heading + 1
Else
Heading = 0
End If
Speed = Speed - 2
End Select
End If
'===
'===<<<WHAT FOLLOWS ARE THE CALCULATIONS FOR SPEED AND ALTITUDE ALTERATION>>>===
'===
Select Case Pitch
Case Is = 1
Speed = Speed - 2
Altitude = Altitude + Speed / 70
imgPitch.Picture = imgPos(1).Picture
Case Is = 2
Speed = Speed - 4
Altitude = Altitude + Speed / 45
imgPitch.Picture = imgPos(0).Picture
Case Is = 3
Speed = Speed - 2
Altitude = Altitude + Speed / 70
imgPitch.Picture = imgPos(7).Picture
Case Is = 4
Speed = Speed - 2
Altitude = Altitude + Speed / 70
imgPitch.Picture = imgPos(6).Picture
Case Is = -1
Speed = Speed + 2
Altitude = Altitude - Speed / 70
imgPitch.Picture = imgPos(5).Picture
Case Is = -2
Speed = Speed + 4
Altitude = Altitude - Speed / 45
imgPitch.Picture = imgPos(4).Picture
Case Is = -3
Speed = Speed + 2
Altitude = Altitude - Speed / 70
imgPitch.Picture = imgPos(3).Picture
Case Is = 0
imgPitch.Picture = imgPos(2).Picture
End Select
Select Case Heading
Case Is = 0
shpRunway.Move shpRunway.Left, shpRunway.Top + Speed
Case Is = 1
shpRunway.Move shpRunway.Left + Speed / 1.41, shpRunway.Top + Speed / 1.41
Case Is = 2
shpRunway.Move shpRunway.Left + Speed, shpRunway.Top
Case Is = 3
shpRunway.Move shpRunway.Left + Speed / 1.41, shpRunway.Top - Speed / 1.41
Case Is = 4
shpRunway.Move shpRunway.Left, shpRunway.Top - Speed
Case Is = 5
shpRunway.Move shpRunway.Left - Speed / 1.41, shpRunway.Top + Speed / 1.41
Case Is = 6
shpRunway.Move shpRunway.Left - Speed, shpRunway.Top
Case Is = 7
shpRunway.Move shpRunway.Left - Speed / 1.41, shpRunway.Top + Speed
End Select
If Pitch = 4 And Roll = 4 Then
Pitch = 0
Roll = 0
End If
If Pitch Then
If Altitude < 0 Then Altitude = 0
lblSpeed.Caption = Int(Speed)
lblAltitude.Caption = Int(Altitude)
lblThrottle.Caption = Throttle
End If
End Sub

Private Sub txtKeys_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 104
ElevatorDown = True
Case Is = 102
AileronRight = True
Case Is = 98
ElevatorUp = True
Case Is = 100
AileronLeft = True
Case Is = 107
ThrottleUp = True
Case Is = 109
ThrottleDown = True
End Select
'MsgBox KeyCode
End Sub

Private Sub txtKeys_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case Is = 104
ElevatorDown = False
Case Is = 102
AileronRight = False
Case Is = 98
ElevatorUp = False
Case Is = 100
AileronLeft = False
Case Is = 107
ThrottleUp = False
Case Is = 109
ThrottleDown = False
End Select

End Sub
