VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Invasion!!!"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   0  'User
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7320
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      TabIndex        =   3
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdReload 
      Caption         =   "Reload"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picScreen 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   120
      ScaleHeight     =   400
      ScaleMode       =   0  'User
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   120
      Width           =   9975
   End
   Begin VB.Label lblKills 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label lblHealth 
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   5
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=====================================================
'HI SCORE FOR THIS GAME: 122025 BY BEN ON JUL 7, 2006
'=====================================================
Private Sub GameLoop()
Dim CurrentTick As Long
Dim LastTick As Long
Const FrameDifference As Long = 10
Const FPS = 100
Me.Show
FrameCount = 0
Randomize
Do
    If Terminated = True Then
        'EndIt
        Exit Do
    End If
    
    CurrentTick = GetTickCount()
       
    If CurrentTick - LastTick > FrameDifference Then
        If Paused = False Then UpdateObjects
        BlitObjects
        'Me.Refresh
        LastTick = CurrentTick
        DoEvents
    Else
        DoEvents
        'Sleep 2
    End If
Loop
End Sub

Private Sub cmdExit_Click()
Terminated = True
End Sub

Private Sub cmdPause_Click()
Select Case Paused
Case Is = True
Paused = False
cmdExit.Visible = False
cmdPause.Caption = "Pause"
Case Is = False
Paused = True
cmdExit.Visible = True
cmdPause.Caption = "Resume"
End Select
End Sub

Private Sub cmdReload_Click()
Call SpawnEnemy
End Sub

Private Sub Form_Load()
Dim i
'load backbuffer into memory
BackgroundDC = GenerateDC(App.Path & "\Graphics\Background.bmp")
BackBuffDC = GenerateDC(App.Path & "\Graphics\Background.bmp")
'load shot graphics
ShotDC = GenerateDC(App.Path & "\Graphics\Shot.bmp")
ShotMaskDC = GenerateDC(App.Path & "\Graphics\ShotMask.bmp")
'load enemy graphics
For i = 1 To 8
EnemyDC(i) = GenerateDC(App.Path & "\Graphics\Enemy" & i & ".bmp")
EnemyMaskDC(i) = GenerateDC(App.Path & "\Graphics\EnemyMask" & i & ".bmp")
Next i
'set variables
Player.Ready = True
Player.Damage = 5
Player.Health = 1000
Terminated = False
Paused = False
ShotMin = 1
ShotCount = 0
EnemyMin = 1
EnemyCount = 0
SpawnThreshold = 1.5
Call InitializeData
Call GameLoop
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Player.TargetX = x
Player.TargetY = y
End Sub

Private Sub picScreen_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Player.DrawX = x
Player.DrawY = y
'If Player.Ready = True Then
Call PlayerFire
Player.Points = Player.Points - 5
If Player.Points < 0 Then Player.Points = 0
'Player.Ready = False
'End If
End Sub
