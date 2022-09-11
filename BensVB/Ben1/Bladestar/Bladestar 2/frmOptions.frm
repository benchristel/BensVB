VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bladestar Options"
   ClientHeight    =   4560
   ClientLeft      =   7350
   ClientTop       =   9330
   ClientWidth     =   4470
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4470
   Begin VB.CheckBox chkRandom 
      Caption         =   "Random Weapon Distribution is OFF"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2895
   End
   Begin VB.HScrollBar scrRespawnTime 
      Height          =   375
      Left            =   120
      Max             =   5
      TabIndex        =   4
      Top             =   1680
      Value           =   4
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit Options"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CommandButton cmdClearScores 
      Caption         =   "Clear Player Scores"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdChangeScoring 
      Caption         =   "Change Scoring Method"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblRespawnTime 
      Caption         =   "Player respawn time is currently 4 seconds."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Label lblScoreStatus 
      Caption         =   "Scoring method is currently wins/losses."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkRandom_Click()
Select Case chkRandom.Value
Case Is = 0
    chkRandom.Caption = "Random Weapon Distribution is OFF"
    RandomWeapons = False
Case Is = 1
    chkRandom.Caption = "Random Weapon Distribution is ON"
    RandomWeapons = True
End Select
End Sub

Private Sub cmdChangeScoring_Click()
Dim i
If MsgBox("Changing scoring settings will clear all scores!" & vbLf & "Are you sure you want to proceed?", vbYesNo, "Change Scoring") = vbYes Then
Select Case ScoreMethod
Case Is = 1
    ScoreMethod = 2
    lblScoreStatus.Caption = "Scoring method is currently kills/deaths."
Case Is = 2
    ScoreMethod = 1
    lblScoreStatus.Caption = "Scoring method is currently wins/losses."
End Select
For i = 0 To 9
PlayerRecord(i).Losses = 0
PlayerRecord(i).Wins = 0
Next i
Call UpdateLabels
End If
End Sub

Private Sub cmdClearScores_Click()
Dim i
For i = 0 To 9
PlayerRecord(i).Losses = 0
PlayerRecord(i).Wins = 0
Next i
Call UpdateLabels
End Sub

Private Sub Command1_Click()
Unload frmOptions
End Sub

Private Sub Form_Load()
Select Case ScoreMethod
Case Is = 1
lblScoreStatus.Caption = "Scoring method is currently wins/losses."
Case Is = 2
lblScoreStatus.Caption = "Scoring method is currently kills/deaths."
End Select
If RandomWeapons = True Then
    chkRandom.Value = 1
Else
    chkRandom.Value = 0
End If
scrRespawnTime.Value = PlayerRespawnTime
End Sub

Private Sub scrRespawnTime_Change()
PlayerRespawnTime = Int(scrRespawnTime.Value)
lblRespawnTime.Caption = "Player respawn time is currently " & PlayerRespawnTime & " seconds."
If PlayerRespawnTime = 1 Then lblRespawnTime.Caption = "Player respawn time is currently 1 second."
End Sub
