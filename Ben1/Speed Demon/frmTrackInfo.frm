VERSION 5.00
Begin VB.Form frmTrackInfo 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Track Information"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmTrackInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Different Track"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   555
      Left            =   60
      TabIndex        =   2
      Top             =   2580
      Width           =   4575
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select This Track"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   2040
      Width           =   4575
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   4515
   End
End
Attribute VB_Name = "frmTrackInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
frmTrackInfo.Visible = True
Select Case Track
Case Is = 0
lblInfo.Caption = "The Grass Track is an easy track intended for beginners." & _
"It doesn't have any special hazards or powerups, so it's great for people who don't want an extremely complex game." & _
"An average first place time would be about 150 seconds.  You must complete 3 two-screen laps to finish."
Case Is = 1
lblInfo.Caption = "Desert Storm is an easy track intended for beginners who want a challenge." & _
"It doesn't have any powerups, but watch out for duststorms!  If a storm springs up, you'll be temporarily blinded." & _
"An average first place time would be just over 160 seconds.  You must complete 3 two-screen laps to finish."
End Select
End Sub

Private Sub lblCancel_Click()
Unload frmTrackInfo
End Sub

Private Sub lblSelect_Click()
Select Case Track
Case Is = 0
Load frmTrack1
Case Is = 1
Load frmTrack2
End Select
Unload frmSelectTrack
Unload frmTrackInfo
End Sub
