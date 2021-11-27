VERSION 5.00
Begin VB.Form frmSelectTrack 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Race Selector"
   ClientHeight    =   11400
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   15240
   ForeColor       =   &H00808080&
   Icon            =   "frmSelectTrack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11400
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Race!"
      BeginProperty Font 
         Name            =   "Beanie"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13620
      TabIndex        =   4
      Top             =   10380
      Width           =   1515
   End
   Begin VB.Label lblRank 
      BackStyle       =   0  'Transparent
      Caption         =   "Your rank is"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   795
      Left            =   3960
      TabIndex        =   15
      Top             =   1500
      Width           =   7515
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Desert Storm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   12840
      TabIndex        =   14
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Image imgTrack 
      Enabled         =   0   'False
      Height          =   1140
      Index           =   1
      Left            =   11580
      Picture         =   "frmSelectTrack.frx":030A
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Grass Track"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   12840
      TabIndex        =   13
      Top             =   780
      Width           =   2055
   End
   Begin VB.Image imgTrack 
      Enabled         =   0   'False
      Height          =   1140
      Index           =   0
      Left            =   11580
      Picture         =   "frmSelectTrack.frx":0FD4
      Stretch         =   -1  'True
      Top             =   420
      Width           =   1140
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "Your score is "
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   675
      Left            =   3960
      TabIndex        =   12
      Top             =   780
      Width           =   7215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "4.6.130 - 1000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   11
      Top             =   5340
      Width           =   2475
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "5.8.120 - 1000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   10
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "4.6.110 - 800000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   9
      Top             =   3060
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Skywriter 1550"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   4
      Left            =   1080
      TabIndex        =   8
      Top             =   4680
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enterprise Equator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   3
      Left            =   1080
      TabIndex        =   7
      Top             =   3540
      Width           =   2475
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enterprise Eclipse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   6
      Top             =   2400
      Width           =   2475
   End
   Begin VB.Image imgCar 
      Enabled         =   0   'False
      Height          =   1020
      Index           =   4
      Left            =   120
      Picture         =   "frmSelectTrack.frx":1C9E
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1020
   End
   Begin VB.Image imgCar 
      Enabled         =   0   'False
      Height          =   1020
      Index           =   3
      Left            =   120
      Picture         =   "frmSelectTrack.frx":1FA8
      Stretch         =   -1  'True
      Top             =   3540
      Width           =   1020
   End
   Begin VB.Image imgCar 
      Enabled         =   0   'False
      Height          =   1020
      Index           =   2
      Left            =   120
      Picture         =   "frmSelectTrack.frx":22B2
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1020
   End
   Begin VB.Label lblResult 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a Vehicle."
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   360
      TabIndex        =   5
      Top             =   5760
      Width           =   12375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "3.4.110 - 300000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   2355
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Skywriter 1200 Xtra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   1260
      Width           =   2535
   End
   Begin VB.Image imgCar 
      Enabled         =   0   'False
      Height          =   1020
      Index           =   1
      Left            =   120
      Picture         =   "frmSelectTrack.frx":25BC
      Stretch         =   -1  'True
      Top             =   1260
      Width           =   1020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "4.6.100 - 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   780
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Skywriter 1200"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1995
   End
   Begin VB.Image imgCar 
      Enabled         =   0   'False
      Height          =   1020
      Index           =   0
      Left            =   120
      Picture         =   "frmSelectTrack.frx":28C6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1020
   End
   Begin VB.Shape shpTrack 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   1155
      Index           =   0
      Left            =   11580
      Top             =   420
      Width           =   1155
   End
   Begin VB.Shape shpTrack 
      BorderColor     =   &H0000C0C0&
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   1155
      Index           =   1
      Left            =   11580
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Game"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Game"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmSelectTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
Load frmTrack1
Unload frmSelectTrack
End Sub

Private Sub Form_Load()
Dim i
'Open "C:\Scores.Dat" For Random As #1 Len = Len(Players)
'
' Initialize File
'
'Players.Name = ""
'Players.Score = 0
'Dim Index As Integer
'For Index = 1 To 1000
'Put #1, Index, Players
'Next Index
'
' End Initialize
'
TotalScore = 500000
frmSelectTrack.Visible = True
Select Case TotalScore
Case Is > 2500000
For i = 0 To 4
imgCar(i).Enabled = True
Next i
Rank = "Speed Demon."
Case Is > 1500000
For i = 0 To 4
imgCar(i).Enabled = True
Next i
Rank = "Driving Ace."
Case Is > 1000000
For i = 0 To 4
imgCar(i).Enabled = True
Next i
Rank = "Hot Rodder."
Case Is > 800000
For i = 0 To 2
imgCar(i).Enabled = True
Next i
Rank = "Expert Driver."
Case Is > 500000
For i = 0 To 1
imgCar(i).Enabled = True
Next i
Rank = "Intermediate Driver."
Case Is > 300000
For i = 0 To 1
imgCar(i).Enabled = True
Next i
Rank = "Beginner Driver."
Case Else
imgCar(0).Enabled = True
Rank = "Rookie."
End Select
lblScore.Caption = "Your Total Score is " & TotalScore
lblRank.Caption = "Your Rank is " & Rank
'If TotalScore = "" Then lblScore.Caption = "Your Total Score is 0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #1
End Sub

Sub imgCar_Click(Index As Integer)
Dim i
For i = 0 To 1
imgTrack(i).Enabled = True
Next i
Select Case Index
Case Is = 0
MaxSpeed = 100
AccelRate = 4
BrakeRate = 6
lblResult.Caption = "Skywriter 1200 -- Choose a Course."
Case Is = 1
MaxSpeed = 110
AccelRate = 3
BrakeRate = 4
lblResult.Caption = "Skywriter 1200 Xtra -- Choose a Course."
Case Is = 2
MaxSpeed = 110
AccelRate = 4
BrakeRate = 6
lblResult.Caption = "Enterprise Eclipse -- Choose a Course."
Case Is = 3
MaxSpeed = 120
AccelRate = 5
BrakeRate = 8
lblResult.Caption = "Enterprise Equator -- Choose a Course."
Case Is = 4
MaxSpeed = 130
AccelRate = 4
BrakeRate = 6
lblResult.Caption = "Skywriter 1550 -- Choose a Course."

 End Select
 End Sub

Private Sub imgTrack_Click(Index As Integer)
Track = Index
Load frmTrackInfo
'Select Case Index
'Case Is = 0
'frmTrackInfo.lblInfo.Caption = "The Grass Track is an easy course intended for beginners.  Without any special hazards or power-ups, this track is great for those who don't want extra complications.  An average first-place time for this track would be about 150 seconds."
'End Select
End Sub

Private Sub mnuLoad_Click()
Dim Player
Player = InputBox("Enter Your Name.")
ScoreNumber = 0
Do
ScoreNumber = ScoreNumber + 1
Get #1, ScoreNumber, Players
        If Trim(Players.Name) = "" Then
        ' A new player
        Players.Name = Player
        Players.Score = 0
        Put #1, ScoreNumber, Players
        Exit Do
        End If
    If Trim(Players.Name) = Player Then
    TotalScore = Players.Score
    Exit Do
    End If
Loop
End Sub

'Private Sub mnuNew_Click()
'Dim i, temp, Name, Password
''Players = Players + 1
'Name = InputBox("Enter your name.", [New Player])
'Password = InputBox("Enter your password.", [New Player])
'TotalScore = 0
'lblScore.Caption = "Your Total Score is 0"
'End Sub

Private Sub mnuSave_Click()
Dim i
    Open "C:\Scores.dat" For Output As #1
    Print #1, TotalScore
    Close #1
End Sub
