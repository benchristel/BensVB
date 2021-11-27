VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar scrResponseOffset 
      Height          =   375
      Left            =   120
      Max             =   30
      Min             =   -30
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.HScrollBar scrVisualOffset 
      Height          =   375
      Left            =   120
      Max             =   30
      Min             =   -30
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.CheckBox chkVisuals 
      Alignment       =   1  'Right Justify
      Caption         =   "Video Effects"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox chkSounds 
      Alignment       =   1  'Right Justify
      Caption         =   "Sound Effects"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.Label lblResponseOffsetCaption 
      Caption         =   "Response Time Adjust: 0 pix"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label lblVisualOffsetCaption 
      Caption         =   "Note Timing Adjust: 0 pix"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkSounds_Click()
If chkSounds.Value = 0 Then
    Player(ActivePlayer).SoundEffects = False
Else
    Player(ActivePlayer).SoundEffects = True
End If
End Sub

Private Sub chkVisuals_Click()
If chkVisuals.Value = 0 Then
    Player(ActivePlayer).VisualEffects = False
Else
    Player(ActivePlayer).VisualEffects = True
End If
End Sub

Private Sub Form_Load()
If Player(ActivePlayer).SoundEffects = True Then
    chkSounds.Value = 1
Else
    chkSounds.Value = 0
End If
If Player(ActivePlayer).VisualEffects = True Then
    chkVisuals.Value = 1
Else
    chkVisuals.Value = 0
End If
scrVisualOffset.Value = Player(ActivePlayer).VisualOffset
scrResponseOffset.Value = Player(ActivePlayer).ResponseOffset
lblVisualOffsetCaption.Caption = "Note Timing Adjust: " & Player(ActivePlayer).VisualOffset & " pix"
lblResponseOffsetCaption.Caption = "Response Timing Adjust: " & Player(ActivePlayer).ResponseOffset & " pix"

End Sub

Private Sub scrResponseOffset_Change()
Player(ActivePlayer).ResponseOffset = scrResponseOffset.Value
lblResponseOffsetCaption.Caption = "Response Timing Adjust: " & Player(ActivePlayer).ResponseOffset & " pix"

End Sub

Private Sub scrVisualOffset_Change()
Player(ActivePlayer).VisualOffset = scrVisualOffset.Value
lblVisualOffsetCaption.Caption = "Note Timing Adjust: " & Player(ActivePlayer).VisualOffset & " pix"
End Sub
