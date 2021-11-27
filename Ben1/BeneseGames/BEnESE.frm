VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000017&
   Caption         =   "Form1"
   ClientHeight    =   9300
   ClientLeft      =   2595
   ClientTop       =   1470
   ClientWidth     =   14955
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   14955
   Begin VB.Timer Timer4 
      Left            =   7560
      Top             =   6240
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   7560
      Top             =   6720
   End
   Begin VB.Timer Timer2 
      Interval        =   6000
      Left            =   7560
      Top             =   7200
   End
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   7560
      Top             =   7680
   End
   Begin VB.Label lblS6 
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      TabIndex        =   9
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label lblS5 
      Caption         =   "Label3"
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label lblS4 
      Caption         =   "Label3"
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label lblS3 
      Caption         =   "Labe"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label lblS2 
      Caption         =   "Label4"
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label lblS1 
      Caption         =   "Label3"
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   7080
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   14760
      Picture         =   "BEnESE.frx":0000
      Stretch         =   -1  'True
      Top             =   -960
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000017&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   5160
      TabIndex        =   1
      Top             =   6240
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000017&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   90
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   3480
      TabIndex        =   0
      Top             =   4080
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Load()
Label1.ForeColor = lblS1.ForeColor
Label2.ForeColor = lblS6.ForeColor
End Sub

Private Sub Timer1_Timer()
Image1.Top = Image1.Top + 90
Image1.Left = Image1.Left - 180
End Sub

Private Sub Timer2_Timer()
Label1.Caption = "BENESE"
Label2.Caption = "GAMES"
Timer2.Enabled = False
Timer1.Enabled = False
Timer3.Enabled = True
Timer4.Enabled = True
End Sub

Private Sub Timer3_Timer()
Label1.Caption = "Moles"
Label2.Caption = ""
cmdPlay.Visible = True
cmdQuit.Visible = True
End Sub

