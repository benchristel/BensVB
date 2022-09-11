VERSION 5.00
Object = "{C61830C1-8B47-11D4-9F3F-0000B45C4CF6}#1.0#0"; "EasySound.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   12675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   845
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9855
      Left            =   240
      ScaleHeight     =   657.908
      ScaleMode       =   0  'User
      ScaleWidth      =   309.367
      TabIndex        =   0
      Top             =   120
      Width           =   4650
   End
   Begin EASYSOUNDLibCtl.ESound ESound 
      Left            =   5040
      OleObjectBlob   =   "frmMain.frx":0000
      Top             =   120
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "= Quit ="
      BeginProperty Font 
         Name            =   "Beanie"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   10560
      Width           =   4650
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   10080
      Width           =   4650
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const FrameDifference As Long = 30
Private Sub Form_Load()
Dim rt As Long, Musicpath As String, i As Integer

ESound.Window = frmMain.hWnd
rt = ESound.InitializeSound()
picScreen.Left = frmMain.ScaleWidth / 2 - 155
lblExit.Left = frmMain.ScaleWidth / 2 - 155
lblScore.Left = frmMain.ScaleWidth / 2 - 155
NoteDC = GenerateDC(App.Path & "\Note.bmp")
NoteMaskDC = GenerateDC(App.Path & "\NoteMask.bmp")
HoleDC = GenerateDC(App.Path & "\Hole.bmp")
HoleLightDC = GenerateDC(App.Path & "\HoleSelected.bmp")
HoleLightMaskDC = GenerateDC(App.Path & "\HoleSelectedMask.bmp")
HoleMaskDC = GenerateDC(App.Path & "\HoleMask.bmp")
BackgroundDC = GenerateDC(App.Path & "\Background.bmp")
BackgroundImageDC = GenerateDC(App.Path & "\Background.bmp")
BackgroundImage2DC = GenerateDC(App.Path & "\Background2.bmp")
TailDC = GenerateDC(App.Path & "\Tail.bmp")
For i = 1 To 3
SpinDC(i) = GenerateDC(App.Path & "\Spin" & i & ".bmp")
SpinMaskDC(i) = GenerateDC(App.Path & "\Spin" & i & "Mask.bmp")
Next i
For i = 1 To 5
SpinState(i) = 0
Next i
'Constants init
ScreenHeight = picScreen.ScaleHeight
'Variables init
Score = 0
AddScore = 0
CumScore = 0
RunBonus = 1
Run = 0
NoteMin = 1
NoteMax = 1
ScrollY = 0
Call LoadSong(LevelFileName)
Musicpath = App.Path & "\Music Files\" & SongFileName & ".wav"
SoundInMemory = ESound.CreateStreamingSound(Musicpath)
Bloop(1) = ESound.CreateStaticSound(App.Path & "\Music Files\Bloop1.wav")
ESound.SetStreamingVolume SoundInMemory, 0
ESound.PlayStreamingSound SoundInMemory, 0
Call GameLoop
End Sub

Private Sub GameLoop()
Dim FrameCount As Integer, i
Dim CurrentTick As Long
Dim LastTick As Long
Static TickDebt As Integer

Me.Show
FrameCount = 0


'sndPlaySound SoundInMemory, &H1 Or &H4 'sound in memory
'WinPlaySound App.Path & "\Music Files\" & SongFileName, True, True
Do
    If Terminated = True Then
        'EndIt
        Exit Do
    End If
    
    CurrentTick = GetTickCount()
    If LastTick = 0 Then GoTo firstframe
    If CurrentTick = LastTick Then GoTo loopthis
        Notespeed = BeatsPerMin / 60 / (1000 / (CurrentTick - LastTick)) * 120
        UpdateObjects
        UpdateKeys
        BlitObjects
        
        
        'Sleep 2
    DoEvents
firstframe:
LastTick = CurrentTick
loopthis:
Loop
ESound.StopStreamingSound SoundInMemory
Unload Me
End Sub

Public Sub BlitObjects()
Dim i As Integer, k As Integer
Dim l As Double, h As Double, s As Double
Const ViewY = 200
Const ViewZ = 500
If RunBonus = 1 Or RunBonus = 2 Then
BitBlt BackgroundDC, 0, 0, 310, 660, BackgroundImageDC, 0, 0, vbSrcCopy
Else
BitBlt BackgroundDC, 0, 0, 310, 660, BackgroundImage2DC, 0, 0, vbSrcCopy
End If
For i = 0 To 4
If KeyPressed(i + 1) = True Then
BitBlt BackgroundDC, i * 60, ScreenHeight - 80, 60, 60, HoleLightMaskDC, 0, 0, vbSrcAnd
BitBlt BackgroundDC, i * 60, ScreenHeight - 80, 60, 60, HoleLightDC, 0, 0, vbSrcPaint
Else
BitBlt BackgroundDC, i * 60, ScreenHeight - 80, 60, 60, HoleMaskDC, 0, 0, vbSrcAnd
BitBlt BackgroundDC, i * 60, ScreenHeight - 80, 60, 60, HoleDC, 0, 0, vbSrcPaint
End If
'BitBlt BackgroundDC, CenterX - 80 + 40 * i, CenterY + 80, 60, 60, HoleMaskDC, 0, 0, vbSrcAnd
'BitBlt BackgroundDC, CenterX - 80 + 40 * i, CenterY + 80, 60, 60, HoleDC, 0, 0, vbSrcPaint
'BitBlt BackgroundDC, CenterX - 120, CenterY - 80 + 40 * i, 60, 60, HoleMaskDC, 0, 0, vbSrcAnd
'BitBlt BackgroundDC, CenterX - 120, CenterY - 80 + 40 * i, 60, 60, HoleDC, 0, 0, vbSrcPaint
'BitBlt BackgroundDC, CenterX + 80, CenterY - 80 + 40 * i, 60, 60, HoleMaskDC, 0, 0, vbSrcAnd
'BitBlt BackgroundDC, CenterX + 80, CenterY - 80 + 40 * i, 60, 60, HoleDC, 0, 0, vbSrcPaint
Next i
For i = NoteMin To NoteMax
If Note(i).Deleted = False Then
l = ViewY * 60 * (Note(i).XOffset - 2) / (Note(i).Distance - ScrollY + ViewY)
h = ViewZ * (Note(i).Distance - ScrollY) / (ViewY + Note(i).Distance - ScrollY)
s = ViewY * 60 / (ViewY + Note(i).Distance - ScrollY)
If Note(i).Duration > 0 Then
StretchBlt BackgroundDC, 150 + l - s / 20, ScreenHeight - 50 - h - Note(i).Duration, s / 10, Note(i).Duration * s / 60, TailDC, 0, ScreenHeight - Note(i).Duration + ScrollY - Note(i).Distance, 6, Note(i).Duration, vbSrcCopy 'ScreenHeight - Note(i).Distance + ScrollY - 80
End If
StretchBlt BackgroundDC, 150 + l - s / 2, ScreenHeight - 50 - h - s / 2, s, s, NoteMaskDC, 0, 0, 60, 60, vbSrcAnd
StretchBlt BackgroundDC, 150 + l - s / 2, ScreenHeight - 50 - h - s / 2, s, s, NoteDC, 0, 0, 60, 60, vbSrcPaint
End If
Next i
'visual effects
If Player(ActivePlayer).VisualEffects = True Then
For i = 0 To 4
If Int(SpinState(i + 1)) > 0 Then
BitBlt BackgroundDC, i * 60, ScreenHeight - 80, 60, 60, SpinMaskDC(4 - Int(SpinState(i + 1))), 0, 0, vbSrcAnd
BitBlt BackgroundDC, i * 60, ScreenHeight - 80, 60, 60, SpinDC(4 - Int(SpinState(i + 1))), 0, 0, vbSrcPaint
SpinState(i + 1) = SpinState(i + 1) - 0.5
End If
Next i
End If
If RunBonus = 1 Or RunBonus = 4 Then
BitBlt picScreen.hdc, 0, 0, picScreen.ScaleWidth, picScreen.ScaleHeight, BackgroundDC, 0, 0, vbSrcCopy
Else
'BitBlt BackgroundDC, 0, 0, picScreen.ScaleWidth, picScreen.ScaleHeight, BackgroundDC, 0, 0, vbSrcInvert
BitBlt picScreen.hdc, 0, 0, picScreen.ScaleWidth, picScreen.ScaleHeight, BackgroundDC, 0, 0, vbNotSrcCopy
End If
'If lblScore.Caption <> CumScore Then lblScore.Caption = CumScore
If lblScore.Caption <> CumScore Then lblScore.Caption = CumScore
End Sub

'Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'CursorX = Int((x + 20) / 40) - 2
'If CursorX > 12 Then CursorX = 12
'If CursorX < 0 Then CursorX = 0
'End Sub
Private Sub lblExit_Click()
Dim i, k, tempplayer As Player
If Player(ActivePlayer).MaxScore(ActiveSong) < CumScore Then Player(ActivePlayer).MaxScore(ActiveSong) = CumScore
frmStartup.lblSongSelect(ActiveSong - songscroll - 1).Caption = songname(ActiveSong) & " - " & Player(ActivePlayer).MaxScore(ActiveSong)
Player(ActivePlayer).TotalScore = 0
For i = 1 To songcount
Player(ActivePlayer).TotalScore = Player(ActivePlayer).TotalScore + Player(ActivePlayer).MaxScore(i)
Next i
For i = 1 To ActivePlayer - 1
If Player(ActivePlayer).TotalScore > Player(i).TotalScore Then
tempplayer = Player(ActivePlayer)
For k = ActivePlayer - 1 To i Step -1
Player(k + 1) = Player(k)
Next k
Player(i) = tempplayer
ActivePlayer = i
Exit For
End If
Next i
Terminated = True
End Sub

Private Sub picScreen_Click()

End Sub
