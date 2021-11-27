VERSION 5.00
Begin VB.Form frmBPMCalc 
   Caption         =   "BPM Calculator"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblOutput 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Tap Spacebar for beat.  Click here to reset."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmBPMCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim terminated As Boolean
Private Sub Form_Load()
Dim i As Integer
Dim CurrentTick As Long
Dim LastTick As Long
Const FrameDifference As Long = 20
Const FPS = 50
Me.Show
BeatCount = -1
Do
    If terminated = True Then
        'EndIt
        Exit Do
    End If
    
    CurrentTick = GetTickCount()
       
        If (GetKeyState(vbKeySpace) And KEY_PRESSED) And (KeyPressed = False) Then
        If BeatCount = -1 Then
        BeatCount = 0
        GoTo startcounting
        End If
        Update (CurrentTick - LastTick)
startcounting:
        KeyPressed = True
        LastTick = CurrentTick
        End If
        If (GetKeyState(vbKeySpace) And KEY_PRESSED) = False Then KeyPressed = False
        DoEvents
Loop
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
terminated = True
Unload Me
End Sub

Private Sub lblInstructions_Click()
BPM = 0
TimeElapsed = 0
BeatCount = 0
AverageTime = 0
End Sub
