VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form2"
   ScaleHeight     =   5910
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   5760
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Play"
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrColors 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   5640
      Top             =   3360
   End
   Begin VB.Timer TimerE 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   8280
      Top             =   2520
   End
   Begin VB.Timer TimerL 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   7680
      Top             =   2520
   End
   Begin VB.Timer TimerG2 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   7080
      Top             =   2520
   End
   Begin VB.Timer TimerG1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   6360
      Top             =   2520
   End
   Begin VB.Timer TimerI 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   5880
      Top             =   2520
   End
   Begin VB.Timer TimerU 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   5400
      Top             =   2520
   End
   Begin VB.Timer TimerQ 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   4440
      Top             =   2520
   End
   Begin VB.Timer TimerS 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   3360
      Top             =   2520
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Form1.Visible = True
Form2.Visible = False
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
TimerS.Enabled = True
End Sub

Private Sub TimerE_Timer()
Label1.Caption = "Squiggle"
TimerE.Enabled = False
tmrColors.Enabled = True
cmdQuit.Visible = True
cmdOk.Visible = True
End Sub

Private Sub TimerG1_Timer()
Label1.Caption = "Squig"
TimerG1.Enabled = False
TimerG2.Enabled = True

End Sub

Private Sub TimerG2_Timer()
Label1.Caption = "Squigg"
TimerG2.Enabled = False
TimerL.Enabled = True
End Sub

Private Sub TimerI_Timer()
Label1.Caption = "Squi"
TimerI.Enabled = False
TimerG1.Enabled = True
End Sub

Private Sub TimerL_Timer()
Label1.Caption = "Squiggl"
TimerL.Enabled = False
TimerE.Enabled = True
End Sub

Private Sub TimerQ_Timer()
Label1.Caption = "Sq"
TimerQ.Enabled = False
TimerU.Enabled = True
End Sub

Private Sub TimerS_Timer()
Label1.Caption = "S"
TimerS.Enabled = False
TimerQ.Enabled = True
End Sub

Private Sub TimerU_Timer()
Label1.Caption = "Squ"
TimerU.Enabled = False
TimerI.Enabled = True
End Sub

Private Sub tmrColors_Timer()
Label1.BackColor = QBColor(9)
Form2.BackColor = QBColor(9)
End Sub
