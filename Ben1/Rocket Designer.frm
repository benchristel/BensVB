VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Rocket Designer"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraTubes 
      Caption         =   "Tube Dimentions"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   7200
      Width           =   2415
      Begin VB.TextBox txtLength 
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Text            =   "12"
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "lnches long"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Shape shpRightFin 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   11640
      Top             =   7320
      Width           =   735
   End
   Begin VB.Shape shpLeftFin 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H80000007&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   10440
      Top             =   7320
      Width           =   735
   End
   Begin VB.Shape shpTube 
      BorderColor     =   &H80000007&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   6000
      Left            =   11160
      Top             =   2040
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TubeLength, prevTubeLength, TubeDiff

Private Sub Form_Load()
TubeLength = 12
prevTubeLength = 12
End Sub

Private Sub txtLength_Change()
Val (txtLength.Text)
If txtLength.Text < 1 Or txtLength.Text = "" Then
MsgBox "Tube must be at least one inch long.", 48, "Tube Length Error"
Exit Sub
End If
TubeLength = txtLength.Text
TubeLength = TubeLength * 500
prevTubeLength = prevTubeLength * 500
shpTube.Height = TubeLength
TubeDiff = TubeLength - prevTubeLength
TubeDiff = TubeDiff / 500 / 15
shpTube.Move shpTube.Top + TubeDiff
prevTubeLength = txtLength.Text
End Sub
