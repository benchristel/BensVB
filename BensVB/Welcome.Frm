VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome"
   ClientHeight    =   5715
   ClientLeft      =   1860
   ClientTop       =   1815
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "Welcome.frx":0000
   ScaleHeight     =   5715
   ScaleWidth      =   7830
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   600
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   3360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Light!"
      Height          =   375
      Left            =   1080
      MaskColor       =   &H000000FF&
      TabIndex        =   1
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800080&
      BorderStyle     =   5  'Dash-Dot-Dot
      FillColor       =   &H00800080&
      Height          =   1935
      Left            =   480
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Learn Visual Basic Now"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Image imgEmpty 
      Height          =   480
      Left            =   1440
      Picture         =   "Welcome.frx":1D35E
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBolt 
      Height          =   3285
      Left            =   4680
      Picture         =   "Welcome.frx":1D668
      Top             =   0
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Image imgFire 
      Height          =   495
      Left            =   1080
      Picture         =   "Welcome.frx":225F2
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image imgCan 
      Height          =   720
      Left            =   4440
      Picture         =   "Welcome.frx":228FC
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'When the user clicks the Light button...
    imgBolt.Visible = True  'display lightning
    Timer1.Enabled = True   'start fire timer
End Sub

Private Sub Command2_Click()
    End
End Sub


Private Sub cmdLight_Click()
    'When the user clicks the Light button...
    imgBolt.Visible = True  'display lightning
    Timer1.Enabled = True   'start fire timer
End Sub

Private Sub cmdQuit_Click()
    End
End Sub


Private Sub Timer1_Timer()
    imgCan.Picture = imgFire.Picture 'start fire
    imgBolt.Visible = False      'hide lightning
    Timer1.Enabled = False       'stop first timer
    Timer2.Enabled = True        'start delay
End Sub




Private Sub Timer2_Timer()
    'when the burn delay is over...
    imgCan.Picture = imgEmpty.Picture  'empty can
    Timer2.Enabled = False             'stop timer
End Sub


