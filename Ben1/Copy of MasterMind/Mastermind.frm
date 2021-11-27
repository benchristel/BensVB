VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   10080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO"
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
      Left            =   7080
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.Shape shpGuess 
      BackColor       =   &H80000006&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   400
      Index           =   0
      Left            =   3700
      Shape           =   3  'Circle
      Top             =   7200
      Width           =   400
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Drag colors to Peg positions and Click GO"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblOrange 
      BackColor       =   &H000080FF&
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblYellow 
      BackColor       =   &H0000FFFF&
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   6240
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblWhite 
      BackColor       =   &H00FFFFFF&
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   6600
      TabIndex        =   11
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblBrown 
      BackColor       =   &H00404080&
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblBlack 
      BackColor       =   &H00000000&
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   120
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
      Height          =   495
      Index           =   1
      Left            =   4320
      TabIndex        =   5
      Top             =   840
      Width           =   750
   End
   Begin VB.Label lblBlue 
      BackColor       =   &H00FF0000&
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblRed 
      BackColor       =   &H000000FF&
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblGuess 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
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
      Height          =   495
      Index           =   0
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   750
   End
   Begin VB.Label lblGreen 
      BackColor       =   &H0000FF00&
      DragMode        =   1  'Automatic
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      Top             =   120
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
Dim j, Index As Integer
If Trial = 12 Then subexit  'THIS IS THE END OF THE GAME
Trial = Trial + 1
For j = 0 To 4
Index = j + 5 * (Trial - 1)
shpGuess(Index).FillColor = lblGuess(j).BackColor
shpGuess(Index).Visible = True
Next j
End Sub

Private Sub Form_Load()
Dim j As Integer
For j = 0 To 4
lblGuess(j).BackColor = &H8000000F ' grey
Next j
For j = 1 To 59
Load shpGuess(j)
'shpGuess(j).BackColor = &H8000000F ' grey
shpGuess(j).FillColor = RGB(255, 0, 0)
shpGuess(j).Top = 7200 - Int(j / 5) * 480
shpGuess(j).Left = 3700 + (j Mod 5) * 800
If j < 5 Then shpGuess(j).Visible = True
Next j
Trial = 0
End Sub

Private Sub lblGuess_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
lblGuess(Index).BackColor = Source.BackColor
End Sub

