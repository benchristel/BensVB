VERSION 5.00
Begin VB.Form frmFireworks 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctDisplay 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3075
      Left            =   60
      ScaleHeight     =   3075
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      Begin VB.Timer tmrTime 
         Interval        =   100
         Left            =   4020
         Top             =   0
      End
   End
End
Attribute VB_Name = "frmFireworks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExplodeX() As Integer, ExplodeY() As Integer
Dim AngleS, AngleE, Fireworks As Integer
Dim Explode1Len() As Integer, Explode2Len() As Integer
Dim Explode1Time() As Integer, Explode2Time() As Integer
Dim Red1() As Integer, Red2() As Integer, Green1() As Integer
Dim Green2() As Integer, Blue1() As Integer, Blue2() As Integer
Dim pi
Private Sub Form_Load()
    Randomize
    Fireworks = 0
End Sub

Private Sub PctDisplay_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim i
Fireworks = Fireworks + 1
ReDim Red1(1 To Fireworks), Red2(1 To Fireworks), Green1(1 To Fireworks)
ReDim Green2(1 To Fireworks), Blue1(1 To Fireworks), Blue2(1 To Fireworks)
ReDim ExplodeX(1 To Fireworks), ExplodeY(1 To Fireworks)
ReDim Explode1Len(1 To Fireworks), Explode1Time(1 To Fireworks)
pi = Atn(1) * 4
    ExplodeX(Fireworks) = x
    ExplodeY(Fireworks) = Y
    SelectColorPattern
    Explode1Len(Fireworks) = 90
    Explode1Time(Fireworks) = 1
    For i = 1 To 10
    AngleS = i * 36
    AngleE = i * 36
    If AngleS >= 360 Then AngleS = AngleS - 360
    If AngleE >= 360 Then AngleE = AngleE - 360
    AngleS = AngleS * pi / 180
    AngleE = AngleE * pi / 180 + 0.01
    pctDisplay.Circle (ExplodeX(Fireworks), ExplodeY(Fireworks)), Explode1Len(Fireworks), _
    RGB(Red1(Fireworks), Green1(Fireworks), Blue1(Fireworks)), -AngleS, -AngleE
    Next i
    ExplodeTime(Fireworks - 1) = 1
End Sub

Private Sub SelectColorPattern()
Dim color
color = Int(Rnd * 3 + 1)
Select Case color
Case Is = 1
Red1(Fireworks) = 255
Red2(Fireworks) = 150
Green1(Fireworks) = 0
Green2(Fireworks) = 0
Blue1(Fireworks) = 0
Blue2(Fireworks) = 150
Case Is = 2
Red1(Fireworks) = 255
Red2(Fireworks) = 0
Green1(Fireworks) = 0
Green2(Fireworks) = 150
Blue1(Fireworks) = 0
Blue2(Fireworks) = 100
Case Is = 3
Red1(Fireworks) = 180
Red2(Fireworks) = 250
Green1(Fireworks) = 0
Green2(Fireworks) = 0
Blue1(Fireworks) = 255
Blue2(Fireworks) = 20
End Select
End Sub

Private Sub tmrTime_Timer()
Dim i, x
If Fireworks = 0 Then Exit Sub
pctDisplay.Cls
If Fireworks = 2 Then
MsgBox "2"
End If
For i = 1 To Fireworks
Explode1Time(i) = Explode1Time(i) + 1
Explode1Len(i) = Explode1Time(i) * 90
For x = 1 To 10
    AngleS = x * 36
    AngleE = x * 36
    If AngleS >= 360 Then AngleS = AngleS - 360
    If AngleE >= 360 Then AngleE = AngleE - 360
    AngleS = AngleS * pi / 180
    AngleE = AngleE * pi / 180 + 0.01
    pctDisplay.Circle (ExplodeX(i), ExplodeY(i)), Explode1Len(i), _
    RGB(Red1(i), Green1(i), Blue1(i)), -AngleS, -AngleE
Next x
Next i
End Sub
