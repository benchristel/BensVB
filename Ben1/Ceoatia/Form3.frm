VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000007&
   Caption         =   "Form3"
   ClientHeight    =   6795
   ClientLeft      =   120
   ClientTop       =   690
   ClientWidth     =   10680
   LinkTopic       =   "Form3"
   ScaleHeight     =   6795
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMoveLt 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   60
   End
   Begin VB.Timer tmrMoveRt 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   600
      Top             =   60
   End
   Begin VB.TextBox txtNotes 
      Height          =   3795
      Left            =   60
      TabIndex        =   4
      Text            =   "No notes yet..."
      Top             =   840
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label lblQuit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   8040
      TabIndex        =   3
      Top             =   60
      Width           =   2535
   End
   Begin VB.Label lblOptions 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Top             =   60
      Width           =   2535
   End
   Begin VB.Label lblInventory 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   60
      Width           =   2535
   End
   Begin VB.Label lblNotes 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Notes"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2535
   End
   Begin VB.Label lblInvisible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   -60
      TabIndex        =   5
      Top             =   0
      Width           =   10695
   End
   Begin VB.Image ImgSpace 
      Appearance      =   0  'Flat
      Height          =   7200
      Index           =   1
      Left            =   -43200
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   43200
   End
   Begin VB.Image ImgSpace 
      Appearance      =   0  'Flat
      Height          =   7200
      Index           =   0
      Left            =   0
      Picture         =   "Form3.frx":9153
      Stretch         =   -1  'True
      Top             =   0
      Width           =   43200
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label2_Click()

End Sub

Private Sub lblInvisible_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case X
Case Is < 1000
tmrMoveLt.Enabled = True
Case Is > 9800
tmrMoveRt.Enabled = True
Case Else
tmrMoveLt.Enabled = False
tmrMoveRt.Enabled = False
End Select
End Sub

Private Sub lblNotes_Click()
If txtNotes.Visible = True Then
txtNotes.Visible = False
Else
txtNotes.Visible = True
End If
End Sub

Private Sub lblQuit_Click()
Dim result
result = MsgBox("Are you sure you want to quit?  You will lose any changes since your last save.", vbYesNo, "Quit Ceoatia")
If result = 6 Then
End
End If
End Sub

Private Sub tmrMoveLt_Timer()
Dim i
For i = 0 To 1
ImgSpace(i).Left = ImgSpace(i).Left + 200
Next i
Refresh
End Sub

Private Sub tmrMoveRt_Timer()
Dim i
For i = 0 To 1
ImgSpace(i).Left = ImgSpace(i).Left - 200
Refresh
Next i
End Sub
