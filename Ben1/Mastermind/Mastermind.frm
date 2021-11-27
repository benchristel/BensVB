VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MASTERMIND"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboColors 
      Height          =   315
      ItemData        =   "Mastermind.frx":0000
      Left            =   2460
      List            =   "Mastermind.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   720
      Width           =   555
   End
   Begin VB.Frame fraRepeats 
      Caption         =   "Allow:"
      Height          =   735
      Left            =   7080
      TabIndex        =   22
      Top             =   0
      Width           =   1395
      Begin VB.OptionButton optRepeats 
         Caption         =   "Repeats"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton optNoRepeats 
         Caption         =   "No Repeats"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   180
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdPlayAgain 
      Caption         =   "Play Again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9300
      TabIndex        =   21
      Top             =   120
      Width           =   675
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&GO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8580
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblNumberColors 
      Caption         =   "Select Number of Colors:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   27
      Top             =   780
      Width           =   2415
   End
   Begin VB.Label lblSolutions 
      Caption         =   "Solution"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7740
      TabIndex        =   25
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblSolution 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Index           =   4
      Left            =   6840
      TabIndex        =   20
      Top             =   1380
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblSolution 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Index           =   3
      Left            =   6000
      TabIndex        =   19
      Top             =   1380
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblSolution 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Index           =   2
      Left            =   5160
      TabIndex        =   18
      Top             =   1380
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblSolution 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Index           =   1
      Left            =   4320
      TabIndex        =   17
      Top             =   1380
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lblSolution 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Index           =   0
      Left            =   3480
      TabIndex        =   16
      Top             =   1380
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Shape shpGuess 
      BackColor       =   &H80000006&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   0
      Left            =   3700
      Shape           =   3  'Circle
      Top             =   7200
      Width           =   400
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Drag colors to Peg positions and Click GO. Right Click colors to disable/enable"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblOrange 
      BackColor       =   &H000080FF&
      Caption         =   "&o"
      DragMode        =   1  'Automatic
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   13
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblYellow 
      BackColor       =   &H0000FFFF&
      Caption         =   "&y"
      DragMode        =   1  'Automatic
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblWhite 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&w"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   11
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblBrown 
      BackColor       =   &H00404080&
      Caption         =   "&b"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblBlack 
      BackColor       =   &H00000000&
      Caption         =   "&l"
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblGuess 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Peg5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Index           =   4
      Left            =   6840
      TabIndex        =   8
      Top             =   840
      Width           =   750
   End
   Begin VB.Label lblGuess 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Peg4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Index           =   3
      Left            =   6000
      TabIndex        =   7
      Top             =   840
      Width           =   750
   End
   Begin VB.Label lblGuess 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Peg3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Index           =   2
      Left            =   5160
      TabIndex        =   6
      Top             =   840
      Width           =   750
   End
   Begin VB.Label lblGuess 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Peg2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Index           =   1
      Left            =   4320
      TabIndex        =   5
      Top             =   840
      Width           =   750
   End
   Begin VB.Label lblBlue 
      BackColor       =   &H00FF0000&
      Caption         =   "&u"
      DragMode        =   1  'Automatic
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblRed 
      BackColor       =   &H000000FF&
      Caption         =   "&r"
      DragMode        =   1  'Automatic
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblGuess 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Peg1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Index           =   0
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   750
   End
   Begin VB.Label lblGreen 
      BackColor       =   &H0000FF00&
      Caption         =   "&e"
      DragMode        =   1  'Automatic
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   59
      Left            =   360
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   58
      Left            =   840
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   57
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   56
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   55
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   54
      Left            =   360
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   53
      Left            =   840
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   52
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   51
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   50
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   49
      Left            =   360
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   48
      Left            =   840
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   47
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   46
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   45
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   44
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   43
      Left            =   840
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   42
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   41
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   40
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   3360
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   39
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   38
      Left            =   840
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   37
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   36
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   35
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   3840
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   34
      Left            =   360
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   33
      Left            =   840
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   32
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   31
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   30
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   4320
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   29
      Left            =   360
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   28
      Left            =   840
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   27
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   26
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   25
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   4800
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   24
      Left            =   360
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   23
      Left            =   840
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   22
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   21
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   20
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   5280
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   19
      Left            =   360
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   18
      Left            =   840
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   17
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   16
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   15
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   5760
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   14
      Left            =   360
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   13
      Left            =   840
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   12
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   11
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   10
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   6240
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   9
      Left            =   360
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   8
      Left            =   840
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   7
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   6
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   5
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   6720
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   4
      Left            =   360
      Shape           =   3  'Circle
      Top             =   7200
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   3
      Left            =   840
      Shape           =   3  'Circle
      Top             =   7200
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   2
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   7200
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   1
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   7200
      Width           =   200
   End
   Begin VB.Shape shpPegs1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00404040&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   2280
      Shape           =   3  'Circle
      Top             =   7200
      Width           =   200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Trial As Integer


Private Sub cmdGo_Click()
Dim i, j, ScoreIndex, StartIndex, StopIndex  As Integer
Dim Scored_i(4), Scored_j(4), SameGuess As Boolean
'
' Check if user made the exact same guess as last time,
' and if they did, exit the routine.
'
optRepeats.Enabled = False
optNoRepeats.Enabled = False
cboColors.Enabled = False
'
' Check for any grey guesses and warn if you find them
'
For j = 0 To 4
If lblGuess(j).BackColor = &H8000000F Then ' grey
    MsgBox "You haven't made 5 guesses."
    Exit Sub
    End If
Next j
If Trial > 0 Then ' can't check on the very first trial
    SameGuess = True
    For i = 0 To 4
    j = i + 5 * (Trial)
    If shpGuess(j - 5).FillColor <> lblGuess(i).BackColor Then SameGuess = False
    Next i
    If SameGuess = True Then Exit Sub
End If
'
' Transfer the guess to the pegs:
'
Trial = Trial + 1
    For i = 0 To 4
    j = i + 5 * (Trial - 1)
    shpGuess(j).FillColor = lblGuess(i).BackColor
    shpGuess(j).Visible = True
    Next i
'
' Check for Red Scores
'
ScoreIndex = 5 * (Trial - 1)
    For i = 0 To 4
    j = i + 5 * (Trial - 1)
    If shpGuess(j).FillColor = lblSolution(i).BackColor Then
        shpPegs1(ScoreIndex).FillColor = &HFF&
        ScoreIndex = ScoreIndex + 1
        Scored_i(i) = True
        Scored_j(i) = True
        End If
    Next i
    If ScoreIndex = Trial * 5 Then
        For i = 0 To 4
        lblSolution(i).Visible = True
        Next i
        lblSolutions.Visible = True
        MsgBox ("You Won!!!")
        cmdGo.Enabled = False
        cmdPlayAgain.Enabled = True
        optRepeats.Enabled = True
        optNoRepeats.Enabled = True
        cboColors.Enabled = True
        lblSolutions.Visible = True
    End If
'
' Check for White Scores
'
StartIndex = 5 * (Trial - 1)
StopIndex = StartIndex + 4
For i = 0 To 4
    For j = StartIndex To StopIndex
 '       If shpGuess(j).FillColor = lblSolution(i).BackColor And Scored(j Mod 5) = False Then
        If shpGuess(j).FillColor = lblSolution(i).BackColor And Scored_i(i) = False And Scored_j(j Mod 5) = False Then
        shpPegs1(ScoreIndex).FillColor = &HFFFFFF
        ScoreIndex = ScoreIndex + 1
        Scored_i(i) = True
        Scored_j(j Mod 5) = True
        End If
    Next j
Next i
If Trial = 12 Then
    For i = 0 To 4
    lblSolution(i).Visible = True
    Next i
    lblSolutions.Visible = True
    MsgBox "GAME OVER!", 48, ""
    cmdGo.Enabled = False
    cmdPlayAgain.Enabled = True
    optRepeats.Enabled = True
    optNoRepeats.Enabled = True
    cboColors.Enabled = True
    Exit Sub 'THIS IS THE END OF THE GAME
End If
End Sub

Private Sub cmdPlayAgain_Click()
Dim j As Integer
For j = 0 To 59
shpGuess(j).FillColor = &H8000000F ' grey
shpPegs1(j).FillColor = &H8000000F ' grey
shpGuess(j).Visible = False
'
' Make the first five variable visible
'
If j < 5 Then shpGuess(j).Visible = True
Next j
For j = 0 To 4
lblGuess(j).BackColor = &H8000000F ' grey
Next j
Trial = 0
'
lblOrange.Visible = True
lblYellow.Visible = True
lblWhite.Visible = True
For j = 4 To 7
If j + 1 > cboColors.Text Then
    Select Case j
    Case Is = 5:
    lblOrange.Visible = False
    Case Is = 6:
    lblYellow.Visible = False
    Case Is = 7:
    lblWhite.Visible = False
    End Select
End If
Next j
lblSolutions.Visible = False
If optNoRepeats.Value = True Then Call MakeSolution(True)
If optRepeats.Value = True Then Call MakeSolution(False)
cmdGo.Enabled = True
cmdPlayAgain.Enabled = False
cboColors.Enabled = False
optRepeats.Enabled = False
optNoRepeats.Enabled = False
End Sub

Private Sub Form_Load()
Dim Color As Integer
Dim i, j As Integer
'
' Initialize the 'guess' squares:
'
For i = 0 To 4
lblGuess(i).BackColor = &H8000000F ' grey
Next i
'
' Create the array of Guess peg locations
' and set their colors to empty grey
'
For j = 1 To 59
Load shpGuess(j)
shpGuess(j).FillColor = &H8000000F ' grey
shpGuess(j).Top = 7200 - Int(j / 5) * 480
shpGuess(j).Left = 3700 + (j Mod 5) * 800
'
' Make the first five visible
'
If j < 5 Then shpGuess(j).Visible = True
Next j
'
' Set Trial to 0.  This is incremented on each guess
'
Trial = 0
'
' Create the Solution
'
cboColors.ListIndex = 3
cmdGo.Enabled = False
If optNoRepeats.Value = True Then Call MakeSolution(True)
If optRepeats.Value = True Then Call MakeSolution(False)

End Sub

Private Sub lblBlack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If lblBlack.BackColor = &HC8C8C8 Then
    lblBlack.BackColor = &H0
    Else
    lblBlack.BackColor = &HC8C8C8
    End If
End If
End Sub
Private Sub lblBrown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If lblBrown.BackColor = &HC8C8C8 Then
    lblBrown.BackColor = &H404080
    Else
    lblBrown.BackColor = &HC8C8C8
    End If
End If
End Sub
Private Sub lblBlue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If lblBlue.BackColor = &HC8C8C8 Then
    lblBlue.BackColor = &HFF0000
    Else
    lblBlue.BackColor = &HC8C8C8
    End If
End If
End Sub
Private Sub lblgreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If lblGreen.BackColor = &HC8C8C8 Then
    lblGreen.BackColor = &HFF00&
    Else
    lblGreen.BackColor = &HC8C8C8
    End If
End If
End Sub
Private Sub lblred_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If lblRed.BackColor = &HC8C8C8 Then
    lblRed.BackColor = &HFF
    Else
    lblRed.BackColor = &HC8C8C8
    End If
End If
End Sub
Private Sub lblorange_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If lblOrange.BackColor = &HC8C8C8 Then
    lblOrange.BackColor = &H80FF&
    Else
    lblOrange.BackColor = &HC8C8C8
    End If
End If
End Sub
Private Sub lblyellow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If lblYellow.BackColor = &HC8C8C8 Then
    lblYellow.BackColor = &HFFFF&
    Else
    lblYellow.BackColor = &HC8C8C8
    End If
End If
End Sub
Private Sub lblwhite_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If lblWhite.BackColor = &HC8C8C8 Then
    lblWhite.BackColor = &HFFFFFF
    Else
    lblWhite.BackColor = &HC8C8C8
    End If
End If
End Sub
' &H00000000& is black
' &H00404080& is brown
' &H00FF0000& is blue
' &H0000FF00& is green
' &H000000FF& is red
' &H000080FF& is orange
' &H0000FFFF& is yellow
' &H00FFFFFF& is white

Private Sub lblGuess_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
lblGuess(Index).BackColor = Source.BackColor
End Sub


Private Sub MakeSolution(NoRepeats As Boolean)
Dim i, j, Color, NumberColors As Integer
Dim ColorUsed(7) As Boolean
'
'
'
' Create the Solution
'
' &H00000000& is black
' &H00404080& is brown
' &H00FF0000& is blue
' &H0000FF00& is green
' &H000000FF& is red
' &H000080FF& is orange
' &H0000FFFF& is yellow
' &H00FFFFFF& is white
' &H00C0C0C0& is grey (empty)
'
' Initialize the ColorUsed array:
' to make sure there are no repeating colors in the solution.
'
    For j = 0 To 7
    ColorUsed(j) = False
    Next j
'
' Generate the solution
'
Randomize
NumberColors = Val(cboColors.Text)
For i = 0 To 4
NewTry:
Color = Int(Rnd * NumberColors)
    Select Case Color
    Case Is = 0:
    If ColorUsed(Color) = True And NoRepeats Then GoTo NewTry
    lblSolution(i).BackColor = &H0&        ' black
    ColorUsed(Color) = True
    lblSolution(i).Visible = False
    Case Is = 1:
    If ColorUsed(Color) = True And NoRepeats Then GoTo NewTry
    lblSolution(i).BackColor = &H404080          ' brown
    ColorUsed(Color) = True
    lblSolution(i).Visible = False
    Case Is = 2:
    If ColorUsed(Color) = True And NoRepeats Then GoTo NewTry
    lblSolution(i).BackColor = &HFF0000        ' blue
    ColorUsed(Color) = True
    lblSolution(i).Visible = False
    Case Is = 3:
    If ColorUsed(Color) = True And NoRepeats Then GoTo NewTry
    lblSolution(i).BackColor = &HFF00&         ' green
    ColorUsed(Color) = True
    lblSolution(i).Visible = False
    Case Is = 4:
    If ColorUsed(Color) = True And NoRepeats Then GoTo NewTry
    lblSolution(i).BackColor = &HFF&       ' red
    ColorUsed(Color) = True
    lblSolution(i).Visible = False
    Case Is = 5:
    If ColorUsed(Color) = True And NoRepeats Then GoTo NewTry
    lblSolution(i).BackColor = &H80FF&            ' orange
    ColorUsed(Color) = True
    lblSolution(i).Visible = False
    Case Is = 6:
    If ColorUsed(Color) = True And NoRepeats Then GoTo NewTry
    lblSolution(i).BackColor = &HFFFF&           ' yellow
    ColorUsed(Color) = True
    lblSolution(i).Visible = False
    Case Is = 7:
    If ColorUsed(Color) = True And NoRepeats Then GoTo NewTry
    lblSolution(i).BackColor = &HFFFFFF          ' white
    ColorUsed(Color) = True
    lblSolution(i).Visible = False
    End Select
'
' Show solution for debugging:
'
' lblSolution(i).Visible = True
'
'
Next i
End Sub

Private Sub optNoRepeats_Click()
If optNoRepeats.Value = True Then Call MakeSolution(True)
If optRepeats.Value = True Then Call MakeSolution(False)
End Sub

Private Sub optRepeats_Click()
If optNoRepeats.Value = True Then Call MakeSolution(True)
If optRepeats.Value = True Then Call MakeSolution(False)
End Sub
