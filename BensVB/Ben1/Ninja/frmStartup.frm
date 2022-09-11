VERSION 5.00
Begin VB.Form frmStartup 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Ninja Stars"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16320
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   16320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Beanie"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   720
      TabIndex        =   11
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblOptions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Beanie"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblNewPlayer 
      BackStyle       =   0  'Transparent
      Caption         =   "New Player >>"
      BeginProperty Font 
         Name            =   "Beanie"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   4320
      TabIndex        =   9
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Ninja Stars"
      BeginProperty Font 
         Name            =   "Beanie"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   1440
      TabIndex        =   8
      Top             =   6600
      Width           =   11055
   End
   Begin VB.Label lblSelectMode 
      BackColor       =   &H00000000&
      Caption         =   "Select Player"
      BeginProperty Font 
         Name            =   "Beanie"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   4920
      TabIndex        =   7
      Top             =   2040
      Width           =   5055
   End
   Begin VB.Label lblNext 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Next >"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblPrev 
      BackStyle       =   0  'Transparent
      Caption         =   "< Prev"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lblSongSelect 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   4320
      TabIndex        =   4
      Top             =   5160
      Width           =   10215
   End
   Begin VB.Label lblSongSelect 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   4800
      TabIndex        =   3
      Top             =   4680
      Width           =   10215
   End
   Begin VB.Label lblSongSelect 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   5040
      TabIndex        =   2
      Top             =   4200
      Width           =   10215
   End
   Begin VB.Label lblSongSelect 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   4800
      TabIndex        =   1
      Top             =   3720
      Width           =   10215
   End
   Begin VB.Label lblSongSelect 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   4320
      TabIndex        =   0
      Top             =   3240
      Width           =   10215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   3600
      Width           =   2055
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Selectmode As Integer

Private Sub Form_Load()
Dim i As Integer, k As Integer, temp, playersongcount As Integer
Open App.Path & "\data.dat" For Input As #1
Line Input #1, temp
songcount = temp
If songcount = 0 Then
songcount = 1
GoTo closeit
End If
ReDim songname(1 To songcount)
ReDim songfile(1 To songcount)
For i = 1 To songcount
Line Input #1, temp
songfile(i) = temp
Next i
closeit:
Close #1
On Error Resume Next
Open App.Path & "\Players.dat" For Input As #1
Line Input #1, temp
PlayerCount = temp
Line Input #1, temp
playersongcount = temp
ReDim Player(1 To PlayerCount)
For i = 1 To PlayerCount
ReDim Player(i).MaxScore(1 To songcount)
With Player(i)
    Line Input #1, temp
    .Name = temp
    Line Input #1, temp
    .TotalScore = temp
    For k = 1 To playersongcount
    Line Input #1, temp
    .MaxScore(k) = temp
    Next k
    Line Input #1, temp
    .SoundEffects = temp
    Line Input #1, temp
    .VisualEffects = temp
    Line Input #1, temp
    .VisualOffset = temp
    Line Input #1, temp
    .ResponseOffset = temp
End With
Next i
Close #1
For i = 1 To songcount
Open App.Path & "\Level Files\" & songfile(i) & ".dat" For Input As #1
Line Input #1, temp
songname(i) = temp
Close #1
Next i
For i = 0 To 4
If i + 1 + songscroll <= PlayerCount Then lblSongSelect(i).Caption = Player(i + 1 + songscroll).Name & " - " & Player(i + 1 + scngscroll).TotalScore
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer, k As Integer
Open App.Path & "\Players.dat" For Output As #1
Print #1, PlayerCount
Print #1, songcount
For i = 1 To PlayerCount
    Print #1, Player(i).Name
    Print #1, Player(i).TotalScore
    For k = 1 To songcount
    Print #1, Player(i).MaxScore(k)
    Next k
    Print #1, Player(i).SoundEffects
    Print #1, Player(i).VisualEffects
    Print #1, Player(i).VisualOffset
    Print #1, Player(i).ResponseOffset
Next i
Close #1
End Sub

Private Sub lblExit_Click()
Unload Me
End Sub

Private Sub lblNewPlayer_Click()
Dim i
If Selectmode = 1 Then
Selectmode = 0
lblSelectMode.Caption = "Select Player"
lblNewPlayer.Caption = "New Player >>"
lblOptions.Visible = False
songscroll = 0
For i = 0 To 4
    lblSongSelect(i).Enabled = True
    If i + 1 + songscroll <= PlayerCount Then
        lblSongSelect(i).Caption = Player(i + 1).Name & " - " & Player(i + 1).TotalScore
    Else
        lblSongSelect(i).Caption = ""
        lblSongSelect(i).Enabled = False
    End If
Next i
Else
PlayerCount = PlayerCount + 1
ReDim Preserve Player(1 To PlayerCount)
ReDim Preserve Player(PlayerCount).MaxScore(1 To songcount)
With Player(PlayerCount)
For i = 1 To songcount
    .MaxScore(i) = 0
Next i
    .Name = InputBox("Enter your name.", "New Player", "Default" & PlayerCount - 1)
    .TotalScore = 0
    .SoundEffects = True
    .VisualEffects = True
    .ResponseOffset = 0
    .VisualOffset = 0
End With
If Player(PlayerCount).Name = "" Then
    PlayerCount = PlayerCount - 1
    ReDim Preserve Player(1 To PlayerCount)
    Exit Sub
End If
'go to song selection
    ActivePlayer = PlayerCount
    Selectmode = 1
    lblNewPlayer.Caption = "<< Select Player"
    lblSelectMode.Caption = "Select Song"
    lblOptions.Visible = True
    songscroll = 0
    For i = 0 To 4
    lblSongSelect(i).Enabled = True
    If i + 1 + songscroll <= songcount Then
    lblSongSelect(i).Caption = songname(i + 1 + songscroll) & " - " & Player(ActivePlayer).MaxScore(i + 1 + songscroll)
    Else
    lblSongSelect(i).Caption = ""
    lblSongSelect(i).Enabled = False
    End If
    Next i
End If
End Sub

Private Sub lblNext_Click()
Dim i
If songscroll < songcount - 5 Then songscroll = songscroll + 5

UpdateScreen
End Sub

Private Sub lblOptions_Click()
frmOptions.Show
End Sub

Private Sub lblPrev_Click()
Dim i
songscroll = songscroll - 5
If songscroll < 0 Then songscroll = 0

UpdateScreen
End Sub

Private Sub lblSongSelect_Click(Index As Integer)
Dim i
If Selectmode = 1 Then
    ActiveSong = Index + 1 + songscroll
    LevelFileName = songfile(ActiveSong)
    On Error Resume Next
    Load frmMain
    Terminated = False
    UpdateScreen
Else 'player select enabled
    ActivePlayer = Index + 1 + songscroll
    Selectmode = 1
    lblSelectMode.Caption = "Select Song"
    lblNewPlayer.Caption = "<< Select Player"
    lblOptions.Visible = True
    songscroll = 0
    For i = 0 To 4
    lblSongSelect(i).Enabled = True
    If i + 1 + songscroll <= songcount Then
    lblSongSelect(i).Caption = songname(i + 1 + songscroll) & " - " & Player(ActivePlayer).MaxScore(i + 1 + songscroll)
    Else
    lblSongSelect(i).Caption = ""
    lblSongSelect(i).Enabled = False
    End If
    Next i
End If
Player(ActivePlayer).TotalScore = 0
For i = 1 To songcount
Player(ActivePlayer).TotalScore = Player(ActivePlayer).TotalScore + Player(ActivePlayer).MaxScore(i)
Next i
End Sub

Public Sub UpdateScreen()
If Selectmode = 1 Then
For i = 0 To 4
lblSongSelect(i).Enabled = True
If i + 1 + songscroll <= songcount Then
lblSongSelect(i).Caption = songname(i + 1 + songscroll) & " - " & Player(ActivePlayer).MaxScore(i + 1 + songscroll)
Else
lblSongSelect(i).Caption = ""
lblSongSelect(i).Enabled = False
End If
Next i
Else
For i = 0 To 4
lblSongSelect(i).Enabled = True
If i + 1 + songscroll <= PlayerCount Then
lblSongSelect(i).Caption = Player(i + 1 + songscroll).Name & " - " & Player(i + 1 + scngscroll).TotalScore
Else
lblSongSelect(i).Caption = ""
lblSongSelect(i).Enabled = False
End If
Next i
'lblNewPlayer.Caption = "New Player >>"
End If
End Sub
