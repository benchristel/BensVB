VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000FF00&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bushwhacker 0.1"
   ClientHeight    =   4650
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7740
   FillStyle       =   0  'Solid
   Icon            =   "Moles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Timer Moletimer 
      Enabled         =   0   'False
      Index           =   0
      Left            =   3480
      Top             =   2400
   End
   Begin VB.Image imgUp 
      Height          =   750
      Left            =   1800
      Picture         =   "Moles.frx":0442
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblEvents 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   "Click Bush as he appears.  The game stops when you have missed him 10 times. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      TabIndex        =   3
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label lblHits 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hits: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label lblMisses 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Misses: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      TabIndex        =   0
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Image imgMole 
      Height          =   750
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   750
   End
   Begin VB.Image ImgMiss 
      Height          =   750
      Left            =   2640
      Picture         =   "Moles.frx":15604
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgHalf 
      Height          =   750
      Left            =   960
      Picture         =   "Moles.frx":28246
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Image imgHole 
      Height          =   750
      Left            =   120
      Picture         =   "Moles.frx":3AE88
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Begin VB.Menu mnuInstructions 
            Caption         =   "Instructions"
         End
         Begin VB.Menu mnuAbout 
            Caption         =   "About Moles"
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit Game"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Molestate(), Hits, misses, basetime, playing

Private Sub cmdGo_Click()
lblEvents.Caption = ""
Call startgame
End Sub

Private Sub Form_Load()
Dim a
Randomize
On Error Resume Next
For a = 0 To 7
Load Moletimer(a)
Moletimer(a).Enabled = False
Load imgMole(a)
imgMole(a).Stretch = True '  Needed to scale picture of Bush
imgMole(a).Picture = imgHole.Picture

If a < 4 Then
imgMole(a).Move 2200 + 900 * a, 1200
Else
imgMole(a).Move 2200 + 900 * (a - 4), 2100
End If
imgMole(a).Visible = True
Next
On Error GoTo 0
End Sub
Private Sub imgMole_Click(Index As Integer)
If Molestate(Index) = 0 Or Molestate(Index) = 4 Then Exit Sub
If playing = False Then Exit Sub
Beep
Moletimer(Index).Enabled = False
Molestate(Index) = 0
imgMole(Index).Picture = imgHole.Picture
Moletimer(Index).Interval = 500 + basetime * 4 * Rnd
Hits = Hits + 1
lblHits.Caption = "Hits: " & Hits
basetime = basetime * 0.988
Moletimer(Index).Enabled = True
End Sub

Private Sub mnuAbout_Click()
Form2.Show 1
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuInstructions_Click()
Form3.Show 1
End Sub

Private Sub startgame()
Dim a
basetime = 3000
cmdGo.Enabled = False
lblInstructions.Caption = ""
playing = True
Hits = 0
misses = 0
lblHits.Caption = "Hits: 0"
lblMisses.Caption = "Misses: 0"
ReDim Molestate(7)
For a = 0 To 7
imgMole(a).Picture = imgHole.Picture
Moletimer(a).Interval = 500 + basetime * 4 * Rnd
Moletimer(a).Enabled = True
Next
End Sub

Private Sub Moletimer_Timer(Index As Integer)
Moletimer(Index).Enabled = False
If playing = False Then Exit Sub
Select Case Molestate(Index)
Case 0
Molestate(Index) = 1
imgMole(Index).Picture = imgHalf.Picture
Moletimer(Index).Interval = 80
Case 1
Molestate(Index) = 2
imgMole(Index).Picture = imgUp.Picture
Moletimer(Index).Interval = basetime
Case 2
Molestate(Index) = 3
imgMole(Index).Picture = imgHalf.Picture
Moletimer(Index).Interval = 80
Case 3
Molestate(Index) = 4
imgMole(Index).Picture = ImgMiss.Picture
Moletimer(Index).Interval = 250
Call miss
Case 4
Molestate(Index) = 0
imgMole(Index).Picture = imgHole.Picture
Moletimer(Index).Interval = 500 + basetime * 4 * Rnd
End Select
imgMole(Index).Refresh
Moletimer(Index).Enabled = True
End Sub

Private Sub miss()
Dim a
misses = misses + 1
lblMisses.Caption = "misses: " & misses
If misses = 10 Then
lblEvents.Caption = "GAME OVER"
cmdGo.Enabled = True
For a = 0 To 7
Molestate(a) = 0
imgMole(a).Picture = imgHole.Picture
Next
playing = False
End If
End Sub
