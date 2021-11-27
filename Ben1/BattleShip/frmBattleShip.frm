VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRotate 
      Caption         =   "Rotate Ship"
      Height          =   495
      Left            =   8760
      TabIndex        =   5
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton cmdPlaceDestroyer 
      Caption         =   "Destroyer"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton cmdPlaceSub 
      Caption         =   "Submarine"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6540
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlaceCruiser 
      Caption         =   "Cruiser"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton cmdPlaceBattleship 
      Caption         =   "Battleship"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5460
      Width           =   1935
   End
   Begin VB.CommandButton cmdPlaceCarrier 
      Caption         =   "Carrier"
      Height          =   495
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   99
      Left            =   9240
      Picture         =   "frmBattleShip.frx":0000
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   98
      Left            =   8760
      Picture         =   "frmBattleShip.frx":08CA
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   97
      Left            =   8280
      Picture         =   "frmBattleShip.frx":1194
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   96
      Left            =   7800
      Picture         =   "frmBattleShip.frx":1A5E
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   95
      Left            =   7320
      Picture         =   "frmBattleShip.frx":2328
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   94
      Left            =   6840
      Picture         =   "frmBattleShip.frx":2BF2
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   93
      Left            =   6360
      Picture         =   "frmBattleShip.frx":34BC
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   92
      Left            =   5880
      Picture         =   "frmBattleShip.frx":3D86
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   91
      Left            =   5400
      Picture         =   "frmBattleShip.frx":4650
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   90
      Left            =   4920
      Picture         =   "frmBattleShip.frx":4F1A
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   89
      Left            =   9240
      Picture         =   "frmBattleShip.frx":57E4
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   88
      Left            =   8760
      Picture         =   "frmBattleShip.frx":60AE
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   87
      Left            =   8280
      Picture         =   "frmBattleShip.frx":6978
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   86
      Left            =   7800
      Picture         =   "frmBattleShip.frx":7242
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   85
      Left            =   7320
      Picture         =   "frmBattleShip.frx":7B0C
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   84
      Left            =   6840
      Picture         =   "frmBattleShip.frx":83D6
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   83
      Left            =   6360
      Picture         =   "frmBattleShip.frx":8CA0
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   82
      Left            =   5880
      Picture         =   "frmBattleShip.frx":956A
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   81
      Left            =   5400
      Picture         =   "frmBattleShip.frx":9E34
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   80
      Left            =   4920
      Picture         =   "frmBattleShip.frx":A6FE
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   79
      Left            =   9240
      Picture         =   "frmBattleShip.frx":AFC8
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   78
      Left            =   8760
      Picture         =   "frmBattleShip.frx":B892
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   77
      Left            =   8280
      Picture         =   "frmBattleShip.frx":C15C
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   76
      Left            =   7800
      Picture         =   "frmBattleShip.frx":CA26
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   75
      Left            =   7320
      Picture         =   "frmBattleShip.frx":D2F0
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   74
      Left            =   6840
      Picture         =   "frmBattleShip.frx":DBBA
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   73
      Left            =   6360
      Picture         =   "frmBattleShip.frx":E484
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   72
      Left            =   5880
      Picture         =   "frmBattleShip.frx":ED4E
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   71
      Left            =   5400
      Picture         =   "frmBattleShip.frx":F618
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   70
      Left            =   4920
      Picture         =   "frmBattleShip.frx":FEE2
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   69
      Left            =   9240
      Picture         =   "frmBattleShip.frx":107AC
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   68
      Left            =   8760
      Picture         =   "frmBattleShip.frx":11076
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   67
      Left            =   8280
      Picture         =   "frmBattleShip.frx":11940
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   66
      Left            =   7800
      Picture         =   "frmBattleShip.frx":1220A
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   65
      Left            =   7320
      Picture         =   "frmBattleShip.frx":12AD4
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   64
      Left            =   6840
      Picture         =   "frmBattleShip.frx":1339E
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   63
      Left            =   6360
      Picture         =   "frmBattleShip.frx":13C68
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   62
      Left            =   5880
      Picture         =   "frmBattleShip.frx":14532
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   61
      Left            =   5400
      Picture         =   "frmBattleShip.frx":14DFC
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   60
      Left            =   4920
      Picture         =   "frmBattleShip.frx":156C6
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   59
      Left            =   9240
      Picture         =   "frmBattleShip.frx":15F90
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   58
      Left            =   8760
      Picture         =   "frmBattleShip.frx":1685A
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   57
      Left            =   8280
      Picture         =   "frmBattleShip.frx":17124
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   56
      Left            =   7800
      Picture         =   "frmBattleShip.frx":179EE
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   55
      Left            =   7320
      Picture         =   "frmBattleShip.frx":182B8
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   54
      Left            =   6840
      Picture         =   "frmBattleShip.frx":18B82
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   53
      Left            =   6360
      Picture         =   "frmBattleShip.frx":1944C
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   52
      Left            =   5880
      Picture         =   "frmBattleShip.frx":19D16
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   51
      Left            =   5400
      Picture         =   "frmBattleShip.frx":1A5E0
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   50
      Left            =   4920
      Picture         =   "frmBattleShip.frx":1AEAA
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   49
      Left            =   9240
      Picture         =   "frmBattleShip.frx":1B774
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   48
      Left            =   8760
      Picture         =   "frmBattleShip.frx":1C03E
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   47
      Left            =   8280
      Picture         =   "frmBattleShip.frx":1C908
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   46
      Left            =   7800
      Picture         =   "frmBattleShip.frx":1D1D2
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   45
      Left            =   7320
      Picture         =   "frmBattleShip.frx":1DA9C
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   44
      Left            =   6840
      Picture         =   "frmBattleShip.frx":1E366
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   43
      Left            =   6360
      Picture         =   "frmBattleShip.frx":1EC30
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   42
      Left            =   5880
      Picture         =   "frmBattleShip.frx":1F4FA
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   41
      Left            =   5400
      Picture         =   "frmBattleShip.frx":1FDC4
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   40
      Left            =   4920
      Picture         =   "frmBattleShip.frx":2068E
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   39
      Left            =   9240
      Picture         =   "frmBattleShip.frx":20F58
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   38
      Left            =   8760
      Picture         =   "frmBattleShip.frx":21822
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   37
      Left            =   8280
      Picture         =   "frmBattleShip.frx":220EC
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   36
      Left            =   7800
      Picture         =   "frmBattleShip.frx":229B6
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   35
      Left            =   7320
      Picture         =   "frmBattleShip.frx":23280
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   34
      Left            =   6840
      Picture         =   "frmBattleShip.frx":23B4A
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   33
      Left            =   6360
      Picture         =   "frmBattleShip.frx":24414
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   32
      Left            =   5880
      Picture         =   "frmBattleShip.frx":24CDE
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   31
      Left            =   5400
      Picture         =   "frmBattleShip.frx":255A8
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   30
      Left            =   4920
      Picture         =   "frmBattleShip.frx":25E72
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   29
      Left            =   9240
      Picture         =   "frmBattleShip.frx":2673C
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   28
      Left            =   8760
      Picture         =   "frmBattleShip.frx":27006
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   27
      Left            =   8280
      Picture         =   "frmBattleShip.frx":278D0
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   26
      Left            =   7800
      Picture         =   "frmBattleShip.frx":2819A
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   25
      Left            =   7320
      Picture         =   "frmBattleShip.frx":28A64
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   24
      Left            =   6840
      Picture         =   "frmBattleShip.frx":2932E
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   23
      Left            =   6360
      Picture         =   "frmBattleShip.frx":29BF8
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   22
      Left            =   5880
      Picture         =   "frmBattleShip.frx":2A4C2
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   21
      Left            =   5400
      Picture         =   "frmBattleShip.frx":2AD8C
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   20
      Left            =   4920
      Picture         =   "frmBattleShip.frx":2B656
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   19
      Left            =   9240
      Picture         =   "frmBattleShip.frx":2BF20
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   18
      Left            =   8760
      Picture         =   "frmBattleShip.frx":2C7EA
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   17
      Left            =   8280
      Picture         =   "frmBattleShip.frx":2D0B4
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   16
      Left            =   7800
      Picture         =   "frmBattleShip.frx":2D97E
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   15
      Left            =   7320
      Picture         =   "frmBattleShip.frx":2E248
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   14
      Left            =   6840
      Picture         =   "frmBattleShip.frx":2EB12
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   13
      Left            =   6360
      Picture         =   "frmBattleShip.frx":2F3DC
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   12
      Left            =   5880
      Picture         =   "frmBattleShip.frx":2FCA6
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   11
      Left            =   5400
      Picture         =   "frmBattleShip.frx":30570
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   10
      Left            =   4920
      Picture         =   "frmBattleShip.frx":30E3A
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   9
      Left            =   9240
      Picture         =   "frmBattleShip.frx":31704
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   8
      Left            =   8760
      Picture         =   "frmBattleShip.frx":31FCE
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   7
      Left            =   8280
      Picture         =   "frmBattleShip.frx":32898
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   6
      Left            =   7800
      Picture         =   "frmBattleShip.frx":33162
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   5
      Left            =   7320
      Picture         =   "frmBattleShip.frx":33A2C
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   4
      Left            =   6840
      Picture         =   "frmBattleShip.frx":342F6
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   3
      Left            =   6360
      Picture         =   "frmBattleShip.frx":34BC0
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   2
      Left            =   5880
      Picture         =   "frmBattleShip.frx":3548A
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   1
      Left            =   5400
      Picture         =   "frmBattleShip.frx":35D54
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   99
      Left            =   4380
      Picture         =   "frmBattleShip.frx":3661E
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   98
      Left            =   3900
      Picture         =   "frmBattleShip.frx":36EE8
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   97
      Left            =   3420
      Picture         =   "frmBattleShip.frx":377B2
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   96
      Left            =   2940
      Picture         =   "frmBattleShip.frx":3807C
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   95
      Left            =   2460
      Picture         =   "frmBattleShip.frx":38946
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   94
      Left            =   1980
      Picture         =   "frmBattleShip.frx":39210
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   93
      Left            =   1500
      Picture         =   "frmBattleShip.frx":39ADA
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   92
      Left            =   1020
      Picture         =   "frmBattleShip.frx":3A3A4
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   91
      Left            =   540
      Picture         =   "frmBattleShip.frx":3AC6E
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   90
      Left            =   60
      Picture         =   "frmBattleShip.frx":3B538
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   89
      Left            =   4380
      Picture         =   "frmBattleShip.frx":3BE02
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   88
      Left            =   3900
      Picture         =   "frmBattleShip.frx":3C6CC
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   87
      Left            =   3420
      Picture         =   "frmBattleShip.frx":3CF96
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   86
      Left            =   2940
      Picture         =   "frmBattleShip.frx":3D860
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   85
      Left            =   2460
      Picture         =   "frmBattleShip.frx":3E12A
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   84
      Left            =   1980
      Picture         =   "frmBattleShip.frx":3E9F4
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   83
      Left            =   1500
      Picture         =   "frmBattleShip.frx":3F2BE
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   82
      Left            =   1020
      Picture         =   "frmBattleShip.frx":3FB88
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   81
      Left            =   540
      Picture         =   "frmBattleShip.frx":40452
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   80
      Left            =   60
      Picture         =   "frmBattleShip.frx":40D1C
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   79
      Left            =   4380
      Picture         =   "frmBattleShip.frx":415E6
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   78
      Left            =   3900
      Picture         =   "frmBattleShip.frx":41EB0
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   77
      Left            =   3420
      Picture         =   "frmBattleShip.frx":4277A
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   76
      Left            =   2940
      Picture         =   "frmBattleShip.frx":43044
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   75
      Left            =   2460
      Picture         =   "frmBattleShip.frx":4390E
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   74
      Left            =   1980
      Picture         =   "frmBattleShip.frx":441D8
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   73
      Left            =   1500
      Picture         =   "frmBattleShip.frx":44AA2
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   72
      Left            =   1020
      Picture         =   "frmBattleShip.frx":4536C
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   71
      Left            =   540
      Picture         =   "frmBattleShip.frx":45C36
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   70
      Left            =   60
      Picture         =   "frmBattleShip.frx":46500
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   69
      Left            =   4380
      Picture         =   "frmBattleShip.frx":46DCA
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   68
      Left            =   3900
      Picture         =   "frmBattleShip.frx":47694
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   67
      Left            =   3420
      Picture         =   "frmBattleShip.frx":47F5E
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   66
      Left            =   2940
      Picture         =   "frmBattleShip.frx":48828
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   65
      Left            =   2460
      Picture         =   "frmBattleShip.frx":490F2
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   64
      Left            =   1980
      Picture         =   "frmBattleShip.frx":499BC
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   63
      Left            =   1500
      Picture         =   "frmBattleShip.frx":4A286
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   62
      Left            =   1020
      Picture         =   "frmBattleShip.frx":4AB50
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   61
      Left            =   540
      Picture         =   "frmBattleShip.frx":4B41A
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   60
      Left            =   60
      Picture         =   "frmBattleShip.frx":4BCE4
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   59
      Left            =   4380
      Picture         =   "frmBattleShip.frx":4C5AE
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   58
      Left            =   3900
      Picture         =   "frmBattleShip.frx":4CE78
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   57
      Left            =   3420
      Picture         =   "frmBattleShip.frx":4D742
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   56
      Left            =   2940
      Picture         =   "frmBattleShip.frx":4E00C
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   55
      Left            =   2460
      Picture         =   "frmBattleShip.frx":4E8D6
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   54
      Left            =   1980
      Picture         =   "frmBattleShip.frx":4F1A0
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   53
      Left            =   1500
      Picture         =   "frmBattleShip.frx":4FA6A
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   52
      Left            =   1020
      Picture         =   "frmBattleShip.frx":50334
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   51
      Left            =   540
      Picture         =   "frmBattleShip.frx":50BFE
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   50
      Left            =   60
      Picture         =   "frmBattleShip.frx":514C8
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   49
      Left            =   4380
      Picture         =   "frmBattleShip.frx":51D92
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   48
      Left            =   3900
      Picture         =   "frmBattleShip.frx":5265C
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   47
      Left            =   3420
      Picture         =   "frmBattleShip.frx":52F26
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   46
      Left            =   2940
      Picture         =   "frmBattleShip.frx":537F0
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   45
      Left            =   2460
      Picture         =   "frmBattleShip.frx":540BA
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   44
      Left            =   1980
      Picture         =   "frmBattleShip.frx":54984
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   43
      Left            =   1500
      Picture         =   "frmBattleShip.frx":5524E
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   42
      Left            =   1020
      Picture         =   "frmBattleShip.frx":55B18
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   41
      Left            =   540
      Picture         =   "frmBattleShip.frx":563E2
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   40
      Left            =   60
      Picture         =   "frmBattleShip.frx":56CAC
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   39
      Left            =   4380
      Picture         =   "frmBattleShip.frx":57576
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   38
      Left            =   3900
      Picture         =   "frmBattleShip.frx":57E40
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   37
      Left            =   3420
      Picture         =   "frmBattleShip.frx":5870A
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   36
      Left            =   2940
      Picture         =   "frmBattleShip.frx":58FD4
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   35
      Left            =   2460
      Picture         =   "frmBattleShip.frx":5989E
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   34
      Left            =   1980
      Picture         =   "frmBattleShip.frx":5A168
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   33
      Left            =   1500
      Picture         =   "frmBattleShip.frx":5AA32
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   32
      Left            =   1020
      Picture         =   "frmBattleShip.frx":5B2FC
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   31
      Left            =   540
      Picture         =   "frmBattleShip.frx":5BBC6
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   30
      Left            =   60
      Picture         =   "frmBattleShip.frx":5C490
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   29
      Left            =   4380
      Picture         =   "frmBattleShip.frx":5CD5A
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   28
      Left            =   3900
      Picture         =   "frmBattleShip.frx":5D624
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   27
      Left            =   3420
      Picture         =   "frmBattleShip.frx":5DEEE
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   26
      Left            =   2940
      Picture         =   "frmBattleShip.frx":5E7B8
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   25
      Left            =   2460
      Picture         =   "frmBattleShip.frx":5F082
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   24
      Left            =   1980
      Picture         =   "frmBattleShip.frx":5F94C
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   23
      Left            =   1500
      Picture         =   "frmBattleShip.frx":60216
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   22
      Left            =   1020
      Picture         =   "frmBattleShip.frx":60AE0
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   21
      Left            =   540
      Picture         =   "frmBattleShip.frx":613AA
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   20
      Left            =   60
      Picture         =   "frmBattleShip.frx":61C74
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   19
      Left            =   4380
      Picture         =   "frmBattleShip.frx":6253E
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   18
      Left            =   3900
      Picture         =   "frmBattleShip.frx":62E08
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   17
      Left            =   3420
      Picture         =   "frmBattleShip.frx":636D2
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   16
      Left            =   2940
      Picture         =   "frmBattleShip.frx":63F9C
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   15
      Left            =   2460
      Picture         =   "frmBattleShip.frx":64866
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   14
      Left            =   1980
      Picture         =   "frmBattleShip.frx":65130
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   13
      Left            =   1500
      Picture         =   "frmBattleShip.frx":659FA
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   12
      Left            =   1020
      Picture         =   "frmBattleShip.frx":662C4
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   11
      Left            =   540
      Picture         =   "frmBattleShip.frx":66B8E
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   10
      Left            =   60
      Picture         =   "frmBattleShip.frx":67458
      Stretch         =   -1  'True
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   9
      Left            =   4380
      Picture         =   "frmBattleShip.frx":67D22
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   8
      Left            =   3900
      Picture         =   "frmBattleShip.frx":685EC
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   7
      Left            =   3420
      Picture         =   "frmBattleShip.frx":68EB6
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   6
      Left            =   2940
      Picture         =   "frmBattleShip.frx":69780
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   5
      Left            =   2460
      Picture         =   "frmBattleShip.frx":6A04A
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   4
      Left            =   1980
      Picture         =   "frmBattleShip.frx":6A914
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   3
      Left            =   1500
      Picture         =   "frmBattleShip.frx":6B1DE
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   2
      Left            =   1020
      Picture         =   "frmBattleShip.frx":6BAA8
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   1
      Left            =   540
      Picture         =   "frmBattleShip.frx":6C372
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgShip 
      Height          =   480
      Left            =   9060
      Picture         =   "frmBattleShip.frx":6CC3C
      Top             =   6480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgFSquare 
      Height          =   480
      Index           =   0
      Left            =   4920
      Picture         =   "frmBattleShip.frx":6D906
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgSquare 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "frmBattleShip.frx":6E1D0
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShipOnSpace(0 To 99) As String, PlaceShip As String, ShipOnFSpace(0 To 99), PlaceShipLength, PlaceShipPos, PlaceOK As Boolean
Dim fCarrierLoc, fCarrierStatus
Dim fBattleshipLoc, fBattleshipStatus
Dim fCruiserLoc, fCruiserStatus
Dim fSubmarineLoc, fSubmarineStatus
Dim fDestroyerLoc, fDestroyerStatus
Private Sub cmdPlaceBattleship_Click()
PlaceShip = "Battleship"
PlaceShipLength = 4
End Sub

Private Sub cmdPlaceCarrier_Click()
PlaceShip = "Carrier"
PlaceShipLength = 5
End Sub

Private Sub cmdPlaceCruiser_Click()
PlaceShip = "Cruiser"
PlaceShipLength = 3
End Sub

Private Sub cmdPlaceDestroyer_Click()
PlaceShip = "Destroyer"
PlaceShipLength = 2
End Sub

Private Sub cmdPlaceSub_Click()
PlaceShip = "Submarine"
PlaceShipLength = 3
End Sub

Private Sub cmdRotate_Click()
If PlaceShipPos = 1 Then
PlaceShipPos = 2
Else
PlaceShipPos = 1
End If
End Sub

Private Sub Form_Load()
Dim i
Randomize
PlaceShipPos = 1
'
'For i = 1 To 100
'Load imgSquare(i)
'Select Case i
'Case Is < 11
'imgSquare(i).Top = 1
'Case Is < 21
'imgSquare(i).Top = 2
'Case Is < 31
'imgSquare(i).Top = 3
'Case Is < 41
'imgSquare(i).Top = 4
'Case Is < 51
'imgSquare(i).Top = 5
'Case Is < 61
'imgSquare(i).Top = 6
'Case Is < 71
'imgSquare(i).Top = 7
'Case Is < 81
'imgSquare(i).Top = 8
'Case Is < 91
'imgSquare(i).Top = 9
'Case Is < 101
'imgSquare(i).Top = 10
'End Select
''MsgBox imgSquare(i).Top
'With imgSquare(i)
'.Top = imgSquare(i).Top * 480
'.Top = imgSquare(i).Top - 420
'.Left = 60
'.Left = imgSquare(i).Left + Right(i, 1)
'.Left = imgSquare(i).Left * 480
'.Visible = True
'End With
'MsgBox imgSquare(i).Left
'Next i
''y = -480
''For x = 1 To 100
''Load imgFSquare(x)
''If x Mod 10 = 1 Then y = y + 480
''With imgFSquare(x)
''.Left = x Mod 10
''.Left = imgFSquare(x).Left * 480
''.Left = imgFSquare(x).Left + 4920
''.Top = y
''.Top = imgFSquare(x).Top + 60
''.Visible = True
''End With
''Next x
Call GenerateShip("Carrier", 5)
Call GenerateShip("Battleship", 4)
Call GenerateShip("Cruiser", 3)
Call GenerateShip("Submarine", 3)
Call GenerateShip("Destroyer", 2)
PlaceShip = "Carrier"
PlaceShipLength = 5
'lblShipStatus.Caption = "CARRIER AT (, ) HP 0/0" & vbLf _
'& "BATTLESHIP AT (, ) HP 0/0" & vbLf & "CRUISER AT (, ) HP 0/0" & vbLf _
'& "SUBMARINE AT (, ) HP 0/0" & vbLf & "DESTROYER AT (, ) HP 0/0"
End Sub

Private Sub GenerateShip(ShipName As String, ShipLength As Integer)
Dim i, x, y, ShipLoc, ShipPos
redo:
ShipLoc = Int(Rnd * 99)
ShipPos = Int(Rnd * 2 + 1) ' 1 = Horizontal, 2 = vertical
If ShipPos = 1 Then 'if the ship being placed is horizontal then
x = 10 - ShipLength
y = Val(Right(ShipLoc, 1))
If y > x Then 'if the ship sticks off the edge of the grid then
GoTo redo
End If
For i = ShipLoc To ShipLoc + ShipLength - 1 'check if there is a ship in the way
If ShipOnSpace(i) <> "" Then
GoTo redo
End If
Next i
'if you are here the ship was successfully placed
For i = ShipLoc To ShipLoc + ShipLength - 1 'call up the square numbers again to place the ship
ShipOnSpace(i) = ShipName
'imgSquare(i).Picture = imgShip.Picture
Next i
Else 'ship is vertical
x = 10 - ShipLength
y = Val(Left(ShipLoc, 1))
If y > x Then 'if the ship sticks off the edge of the grid then
GoTo redo
End If
For i = ShipLoc To ShipLoc + (ShipLength - 1) * 10 Step 10 'check if there is a ship in the way
If ShipOnSpace(i) <> "" Then
GoTo redo
End If
Next i
'if you are here the ship was successfully placed
For i = ShipLoc To ShipLoc + (ShipLength - 1) * 10 Step 10 'call up the square numbers again to place the ship
ShipOnSpace(i) = ShipName
'imgSquare(i).Picture = imgShip.Picture
Next i
End If
'redo:
'ShipLoc = Int(Rnd * 99)
'ShipPos = Int(Rnd * 2 + 1) ' 1 = Horizontal, 2 = vertical
'If ShipPos = 1 Then
'If Val(Right(ShipLoc, 1)) > 10 - ShipLength Then GoTo redo ' ship won't fit on grid -- try again
'Else
'If Val(Left(ShipLoc, 1)) > 10 - ShipLength Or ShipLoc > 89 Then GoTo redo
'End If
''if you are here the ship was successfully placed
'    If ShipPos = 1 Then
'        For x = ShipLoc To ShipLoc + ShipLength - 1
'        imgSquare(x).Picture = imgShip.Picture
'            If ShipOnSpace(x) = "" Then
'            ShipOnSpace(x) = ShipName
'            Else
'            GoTo redo ' ships overlap
'            End If
'        Next x
'    Else
'        For x = ShipLoc To ShipLoc + (ShipLength * 10) - 1 Step 10
'        On Error GoTo redo
'        imgSquare(x).Picture = imgShip.Picture
'            If ShipOnSpace(x) = "" Then
'            ShipOnSpace(x) = ShipName
'            Else
'            GoTo redo ' ships overlap
'            End If
'        Next x
'    End If
End Sub

Private Sub imgFSquare_Click(Index As Integer)
Dim i, x, y
Static CarrierPlaced, BattleshipPlaced, CruiserPlaced, SubmarinePlaced, DestroyerPlaced
If CarrierPlaced = True And BattleshipPlaced = True And CruiserPlaced = True And SubmarinePlaced = True And _
DestroyerPlaced = True Then Exit Sub
If PlaceShipPos = 1 Then 'if the ship being placed is horizontal then
x = 10 - PlaceShipLength
y = Val(Right(Index, 1))
If y > x Then 'if the ship sticks off the edge of the grid then
MsgBox "You can't put your ship there because it will go off the edge of the grid.", , "Invalid Ship Placement"
Exit Sub
End If
For i = Index To Index + PlaceShipLength - 1 'check if there is a ship in the way
If ShipOnFSpace(i) <> "" Then
MsgBox "You can't put your ship there because it will overlap another ship.", , "Invalid Ship Placement"
Exit Sub
End If
Next i
'if you are here the ship was successfully placed
For i = Index To Index + PlaceShipLength - 1 'call up the square numbers again to place the ship
ShipOnFSpace(i) = PlaceShip
imgFSquare(i).Picture = imgShip.Picture
Next i
Else 'ship is vertical
x = 10 - PlaceShipLength
y = Val(Left(Index, 1))
If y > x Then 'if the ship sticks off the edge of the grid then
MsgBox "You can't put your ship there because it will go off the edge of the grid.", , "Invalid Ship Placement"
Exit Sub
End If
For i = Index To Index + (PlaceShipLength - 1) * 10 Step 10 'check if there is a ship in the way
If ShipOnFSpace(i) <> "" Then
MsgBox "You can't put your ship there because it will overlap another ship.", , "Invalid Ship Placement"
Exit Sub
End If
Next i
'if you are here the ship was successfully placed
For i = Index To Index + (PlaceShipLength - 1) * 10 Step 10 'call up the square numbers again to place the ship
ShipOnFSpace(i) = PlaceShip
imgFSquare(i).Picture = imgShip.Picture
Next i
End If
redo:
Select Case PlaceShip
    Case Is = "Carrier"
PlaceShip = "Battleship"
PlaceShipLength = 4
CarrierPlaced = True
fCarrierLoc = Index
fCarrierStatus = 5
cmdPlaceCarrier.Enabled = False
    Case Is = "Battleship"
PlaceShip = "Cruiser"
PlaceShipLength = 3
BattleshipPlaced = True
fBattleshipLoc = Index
fBattleshipStatus = 4
cmdPlaceBattleship.Enabled = False
    Case Is = "Cruiser"
PlaceShip = "Submarine"
PlaceShipLength = 3
CruiserPlaced = True
fCruiserLoc = Index
fCruiserStatus = 3
cmdPlaceCruiser.Enabled = False
    Case Is = "Submarine"
PlaceShip = "Destroyer"
PlaceShipLength = 2
SubmarinePlaced = True
fSubmarineLoc = Index
fSubmarineStatus = 3
cmdPlaceSub.Enabled = False
    Case Is = "Destroyer"
PlaceShip = "Carrier"
PlaceShipLength = 5
DestroyerPlaced = True
fDestroyerLoc = Index
fDestroyerStatus = 2
cmdPlaceDestroyer.Enabled = False
End Select
if CarrierPlaced = True and
For i = 0 To 99
If ShipOnFSpace(i) = PlaceShip Then GoTo redo
Next i
'lblShipStatus.Caption = "CARRIER AT (" carrierloc) HP 0/0" & vbLf _
'& "BATTLESHIP AT (, ) HP 0/0" & vbLf & "CRUISER AT (, ) HP 0/0" & vbLf _
'& "SUBMARINE AT (, ) HP 0/0" & vbLf & "DESTROYER AT (, ) HP 0/0"
End Sub

Private Sub imgSquare_Click(Index As Integer)
MsgBox ShipOnSpace(Index) & Index
End Sub
