VERSION 5.00
Begin VB.Form frmSetup 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   15240
   ClientLeft      =   0
   ClientTop       =   135
   ClientWidth     =   19080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   15240
   ScaleWidth      =   19080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrCountdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   360
   End
   Begin VB.TextBox txtScoreToWin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   405
      Index           =   2
      Left            =   17880
      TabIndex        =   2
      Text            =   "20"
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtScoreToWin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   405
      Index           =   1
      Left            =   3120
      TabIndex        =   0
      Text            =   "20"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblScrollText 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmStartup.frx":0000
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   41
      Top             =   7560
      Width           =   19095
   End
   Begin VB.Label lblScrollText 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmStartup.frx":00DF
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   40
      Top             =   5040
      Width           =   19095
   End
   Begin VB.Label lblMap 
      BackStyle       =   0  'Transparent
      Caption         =   "Quicksilver"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   2
      Left            =   12120
      TabIndex        =   39
      Top             =   9360
      Width           =   2655
   End
   Begin VB.Label lblVariant 
      BackStyle       =   0  'Transparent
      Caption         =   "Hardcore"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   3
      Left            =   4440
      TabIndex        =   38
      Top             =   9720
      Width           =   2655
   End
   Begin VB.Label lblOptions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "||| OPTIONS |||"
      BeginProperty Font 
         Name            =   "Weltron Urban"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   5160
      TabIndex        =   37
      Top             =   14520
      Width           =   2775
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "||| HELP |||"
      BeginProperty Font 
         Name            =   "Weltron Urban"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   2640
      TabIndex        =   36
      Top             =   14520
      Width           =   2055
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V. 2.0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4440
      TabIndex        =   35
      Top             =   6840
      Width           =   10455
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B L A D E S T A R"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   4320
      TabIndex        =   34
      Top             =   6000
      Width           =   10695
   End
   Begin VB.Label lblMap 
      BackStyle       =   0  'Transparent
      Caption         =   "Giza"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   1
      Left            =   12120
      TabIndex        =   33
      Top             =   9000
      Width           =   2655
   End
   Begin VB.Label lblMap 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Oracle"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   0
      Left            =   12120
      TabIndex        =   32
      Top             =   8640
      Width           =   2655
   End
   Begin VB.Label lblSelectMapTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Map:"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   12120
      TabIndex        =   31
      Top             =   8280
      Width           =   2655
   End
   Begin VB.Label lblVariant 
      BackStyle       =   0  'Transparent
      Caption         =   "Sabers"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   2
      Left            =   4440
      TabIndex        =   30
      Top             =   9360
      Width           =   2655
   End
   Begin VB.Label lblVariant 
      BackStyle       =   0  'Transparent
      Caption         =   "Grenades"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   29
      Top             =   9000
      Width           =   2655
   End
   Begin VB.Label lblVariant 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   28
      Top             =   8640
      Width           =   2655
   End
   Begin VB.Label lblSelectVariantTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Variant:"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   4440
      TabIndex        =   27
      Top             =   8280
      Width           =   2655
   End
   Begin VB.Label lblPlayer2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player10     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   9
      Left            =   15120
      TabIndex        =   26
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label lblPlayer2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player9     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   8
      Left            =   15120
      TabIndex        =   25
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Label lblPlayer2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player8     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   7
      Left            =   15120
      TabIndex        =   24
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label lblPlayer2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player7     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   6
      Left            =   15120
      TabIndex        =   23
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label lblPlayer2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player6     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   5
      Left            =   15120
      TabIndex        =   22
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label lblPlayer2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player5     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   4
      Left            =   15120
      TabIndex        =   21
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label lblPlayer2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player4     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   3
      Left            =   15120
      TabIndex        =   20
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label lblPlayer2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player3     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   2
      Left            =   15120
      TabIndex        =   19
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label lblPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player10     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   9
      Left            =   360
      TabIndex        =   18
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label lblPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player9     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   17
      Top             =   3720
      Width           =   3615
   End
   Begin VB.Label lblPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player8     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   7
      Left            =   360
      TabIndex        =   16
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Label lblPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player7     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   15
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Label lblPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player6     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label lblPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player5     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   13
      Top             =   2280
      Width           =   3615
   End
   Begin VB.Label lblPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player4     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label lblPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player3     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label lblPlayer2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player2     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   1
      Left            =   15120
      TabIndex        =   10
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label lblPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player2     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   9
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label lblPlayer2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player1     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   0
      Left            =   15120
      TabIndex        =   8
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblPlayer1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player1     0/0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label lblCountdown 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   15240
      TabIndex        =   6
      Top             =   14040
      Width           =   3735
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "||| EXIT |||"
      BeginProperty Font 
         Name            =   "Weltron Urban"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   14520
      Width           =   2055
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "||| START GAME |||"
      BeginProperty Font 
         Name            =   "Weltron Urban"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   15120
      TabIndex        =   4
      Top             =   14520
      Width           =   3855
   End
   Begin VB.Label lblScoreToWin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Score To Win (Player 2)"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Index           =   2
      Left            =   14760
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblScoreToWin 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Score To Win (Player 1)"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label lblScrollText 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmStartup.frx":01C0
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   -120
      TabIndex        =   42
      Top             =   4920
      Width           =   19095
   End
   Begin VB.Label lblScrollText 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmStartup.frx":0287
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   43
      Top             =   7620
      Width           =   19095
   End
   Begin VB.Label lblTitle2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   120
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   7440
      TabIndex        =   44
      Top             =   5040
      Width           =   4335
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Countdown
Private Sub Form_Load()
Dim i
'=================
'Set Screen Colors
'=================
Me.BackColor = RGB(10, 30, 10)
For i = 1 To 2
txtScoreToWin(i).BackColor = RGB(10, 30, 10)
txtScoreToWin(i).ForeColor = RGB(10, 140, 70)
lblScoreToWin(i).ForeColor = RGB(20, 200, 100)
Next i
lblTitle.ForeColor = RGB(20, 200, 100)
lblTitle2.ForeColor = RGB(20, 75, 30)
lblSelectMapTitle.ForeColor = RGB(20, 200, 100)
lblSelectVariantTitle.ForeColor = RGB(20, 200, 100)
For i = lblVariant.LBound To lblVariant.UBound
lblVariant(i).ForeColor = RGB(10, 140, 70)
Next i
For i = lblMap.LBound To lblMap.UBound
lblMap(i).ForeColor = RGB(10, 140, 70)
Next i
For i = 0 To 9
lblPlayer1(i).ForeColor = RGB(10, 140, 70)
lblPlayer2(i).ForeColor = RGB(10, 140, 70)
Next i
lblExit.ForeColor = RGB(20, 200, 100)
lblStart.ForeColor = RGB(20, 200, 100)
lblOptions.ForeColor = RGB(20, 200, 100)
lblHelp.ForeColor = RGB(20, 200, 100)
lblScrollText(1).ForeColor = RGB(10, 35, 15)
lblScrollText(2).ForeColor = RGB(10, 35, 15)
lblScrollText(3).ForeColor = RGB(10, 45, 20)
lblScrollText(4).ForeColor = RGB(10, 45, 20)

'==============
' set variables
'==============
ScoreToWin(1) = txtScoreToWin(1).Text
ScoreToWin(2) = txtScoreToWin(2).Text
For i = 0 To 9
With PlayerRecord(i)
    .Losses = 0
    .Wins = 0
    .Name = "Player" & i + 1
End With
Next i
Player(1).RecordIndex = 0
Player(2).RecordIndex = 1
RuleVariant = 0
StartWeapon(1) = 1
StartWeapon(2) = 1
ScoreMethod = 1
PlayerRespawnTime = 4
Call UpdateLabels
If Screen.Height <> 15360 Or Screen.Width <> 19200 Then
MsgBox "Bladestar requires a screen resolution of 1280 by 1040 pixels to display properly." & vbLf & "Please reset your resolution before running Bladestar.", , "Display Error"
Unload Me
Set frmSetup = Nothing
End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblExit_Click()
If MsgBox("Are you sure you want to leave the game?", vbYesNo, "Exit Game") = vbYes Then
Unload Me
Set frmSetup = Nothing
End If
End Sub

Private Sub lblHelp_Click()
frmHelp.Show 1
End Sub

Private Sub lblMap_Click(Index As Integer)
Dim i
MapSelected = Index
For i = 0 To lblMap.UBound
lblMap(i).BorderStyle = 0
Next i
lblMap(Index).BorderStyle = 1
End Sub

Private Sub lblOptions_Click()
frmOptions.Show 1
End Sub

Private Sub lblPlayer1_DblClick(Index As Integer)
On Error Resume Next
PlayerRecord(Index).Name = InputBox("Enter a name for this player", "New Player", "Player" & Index + 1)
lblPlayer1(Index).Caption = PlayerRecord(Index).Name & "     " & PlayerRecord(Index).Wins & "/" & PlayerRecord(Index).Losses
lblPlayer2(Index).Caption = PlayerRecord(Index).Name & "     " & PlayerRecord(Index).Wins & "/" & PlayerRecord(Index).Losses
Player(1).RecordIndex = Index
If Player(2).RecordIndex = Index Then Player(2).RecordIndex = 0
Call UpdateLabels
End Sub
Private Sub lblPlayer1_Click(Index As Integer)
Dim i
lblPlayer2(Player(1).RecordIndex).Enabled = True
For i = 0 To 9
lblPlayer1(i).BorderStyle = 0
If i = Index Then
    lblPlayer2(i).Enabled = False
End If
Next i
lblPlayer1(Index).BorderStyle = 1
Player(1).RecordIndex = Index
End Sub

Private Sub lblPlayer2_Click(Index As Integer)
Dim i
lblPlayer1(Player(2).RecordIndex).Enabled = True
For i = 0 To 9
lblPlayer2(i).BorderStyle = 0
If i = Index Then
    lblPlayer1(i).Enabled = False
End If
Next i
lblPlayer2(Index).BorderStyle = 1
Player(2).RecordIndex = Index
End Sub
Private Sub lblPlayer2_DblClick(Index As Integer)
On Error Resume Next
PlayerRecord(Index).Name = InputBox("Enter a name for this player", "New Player", "Player" & Index + 1)
lblPlayer2(Index).Caption = PlayerRecord(Index).Name & "     " & PlayerRecord(Index).Wins & "/" & PlayerRecord(Index).Losses
lblPlayer1(Index).Caption = PlayerRecord(Index).Name & "     " & PlayerRecord(Index).Wins & "/" & PlayerRecord(Index).Losses
Player(2).RecordIndex = Index
If Player(1).RecordIndex = Index Then Player(1).RecordIndex = 0
Call UpdateLabels
End Sub

Private Sub lblStart_Click()
Countdown = 5
tmrCountdown.Enabled = True
lblCountdown.Caption = "Starting Game in 5..."
End Sub

Private Sub lblVariant_Click(Index As Integer)
Dim i
RuleVariant = Index
For i = 0 To lblVariant.UBound
lblVariant(i).BorderStyle = 0
Next i
lblVariant(Index).BorderStyle = 1
End Sub

Private Sub tmrCountdown_Timer()
Countdown = Countdown - 1
lblCountdown.Caption = "Starting Game in " & Countdown & "..."
If Countdown = 0 Then
tmrCountdown.Enabled = False
lblCountdown.Caption = ""
On Error Resume Next
frmMain.Visible = True
End If
End Sub

Private Sub txtScoreToWin_Change(Index As Integer)
txtScoreToWin(Index).Text = Int(Val(txtScoreToWin(Index).Text))
If txtScoreToWin(Index).Text > 100 Then txtScoreToWin(Index).Text = 100
If txtScoreToWin(Index).Text < 1 Then txtScoreToWin(Index).Text = 1
ScoreToWin(Index) = txtScoreToWin(Index).Text
End Sub

