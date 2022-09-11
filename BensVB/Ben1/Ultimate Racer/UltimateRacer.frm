VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0000C000&
   Caption         =   "Form1"
   ClientHeight    =   10830
   ClientLeft      =   -435
   ClientTop       =   735
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   11400
   Begin VB.Timer tmrPwrDn 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12120
      Top             =   60
   End
   Begin VB.Timer tmrMain 
      Interval        =   100
      Left            =   12600
      Top             =   60
   End
   Begin VB.TextBox txtMove 
      Height          =   315
      Left            =   480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   -1000
      Width           =   195
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Command1"
      Height          =   315
      Left            =   12420
      TabIndex        =   1
      Top             =   -1000
      Width           =   315
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   39
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   38
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   37
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   36
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   35
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   420
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   34
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   33
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   180
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   32
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   31
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   30
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   29
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   28
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   27
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   26
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   25
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   24
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   23
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   22
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   21
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   20
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   19
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   18
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   17
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   16
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   15
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   14
      Left            =   900
      Shape           =   2  'Oval
      Top             =   2340
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   13
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   12
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   11
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   1020
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   10
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   2160
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   8
      Left            =   5520
      Shape           =   2  'Oval
      Top             =   -1000
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   7
      Left            =   5340
      Shape           =   2  'Oval
      Top             =   -1000
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Index           =   6
      Left            =   5520
      Shape           =   2  'Oval
      Top             =   -1000
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   5
      Left            =   5760
      Shape           =   2  'Oval
      Top             =   -1000
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   4
      Left            =   5340
      Shape           =   2  'Oval
      Top             =   -1000
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   5820
      Shape           =   2  'Oval
      Top             =   -1000
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   5520
      Shape           =   2  'Oval
      Top             =   -1000
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   5340
      Shape           =   2  'Oval
      Top             =   -1000
      Width           =   495
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   5520
      Shape           =   2  'Oval
      Top             =   -1000
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   4
      Top             =   4620
      Width           =   15135
   End
   Begin VB.Shape shpSmoke 
      BorderColor     =   &H80000005&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   9
      Left            =   -500
      Shape           =   2  'Oval
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape shpPlyr2 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   1155
      Left            =   8580
      Top             =   5460
      Width           =   615
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   49
      Left            =   13680
      Shape           =   5  'Rounded Square
      Top             =   1140
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   48
      Left            =   10320
      Shape           =   5  'Rounded Square
      Top             =   540
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   47
      Left            =   12300
      Shape           =   5  'Rounded Square
      Top             =   2580
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   46
      Left            =   960
      Shape           =   5  'Rounded Square
      Top             =   840
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   45
      Left            =   2280
      Shape           =   5  'Rounded Square
      Top             =   1440
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   44
      Left            =   7080
      Shape           =   5  'Rounded Square
      Top             =   0
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   43
      Left            =   13980
      Shape           =   5  'Rounded Square
      Top             =   1920
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   42
      Left            =   10800
      Shape           =   5  'Rounded Square
      Top             =   6540
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   41
      Left            =   6900
      Shape           =   5  'Rounded Square
      Top             =   2580
      Width           =   495
   End
   Begin VB.Shape shpDamage 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   13080
      Top             =   60
      Width           =   555
   End
   Begin VB.Shape shpDamage 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   13080
      Top             =   180
      Width           =   555
   End
   Begin VB.Shape shpDamage 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   13080
      Top             =   300
      Width           =   555
   End
   Begin VB.Shape shpDamage 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   13080
      Top             =   420
      Width           =   555
   End
   Begin VB.Shape shpDamage 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   13080
      Top             =   540
      Width           =   555
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   300
      TabIndex        =   3
      Top             =   6540
      Width           =   15135
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   40
      Left            =   2460
      Shape           =   5  'Rounded Square
      Top             =   8220
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   39
      Left            =   3960
      Shape           =   5  'Rounded Square
      Top             =   8940
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   38
      Left            =   2520
      Shape           =   5  'Rounded Square
      Top             =   10020
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   37
      Left            =   780
      Shape           =   5  'Rounded Square
      Top             =   6960
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   36
      Left            =   8340
      Shape           =   5  'Rounded Square
      Top             =   9420
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   35
      Left            =   780
      Shape           =   5  'Rounded Square
      Top             =   9120
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   34
      Left            =   7560
      Shape           =   5  'Rounded Square
      Top             =   7560
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   33
      Left            =   13440
      Shape           =   5  'Rounded Square
      Top             =   10080
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   32
      Left            =   13380
      Shape           =   5  'Rounded Square
      Top             =   8460
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   31
      Left            =   5220
      Shape           =   5  'Rounded Square
      Top             =   6600
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   30
      Left            =   7080
      Shape           =   5  'Rounded Square
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   29
      Left            =   2940
      Shape           =   5  'Rounded Square
      Top             =   4560
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   28
      Left            =   780
      Shape           =   5  'Rounded Square
      Top             =   4140
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   27
      Left            =   5520
      Shape           =   5  'Rounded Square
      Top             =   2640
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   26
      Left            =   5520
      Shape           =   5  'Rounded Square
      Top             =   1200
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   25
      Left            =   10200
      Shape           =   5  'Rounded Square
      Top             =   7440
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   24
      Left            =   12900
      Shape           =   5  'Rounded Square
      Top             =   6000
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   23
      Left            =   7320
      Shape           =   5  'Rounded Square
      Top             =   3600
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   22
      Left            =   13140
      Shape           =   5  'Rounded Square
      Top             =   4020
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   21
      Left            =   4740
      Shape           =   5  'Rounded Square
      Top             =   6120
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   20
      Left            =   840
      Shape           =   5  'Rounded Square
      Top             =   6240
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   19
      Left            =   11580
      Shape           =   5  'Rounded Square
      Top             =   10140
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   18
      Left            =   11880
      Shape           =   5  'Rounded Square
      Top             =   5940
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   17
      Left            =   900
      Shape           =   5  'Rounded Square
      Top             =   1320
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   16
      Left            =   12180
      Shape           =   5  'Rounded Square
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   15
      Left            =   2280
      Shape           =   5  'Rounded Square
      Top             =   8760
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   14
      Left            =   9780
      Shape           =   5  'Rounded Square
      Top             =   7560
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   13
      Left            =   9480
      Shape           =   5  'Rounded Square
      Top             =   3960
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   12
      Left            =   3660
      Shape           =   5  'Rounded Square
      Top             =   7440
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   11
      Left            =   6420
      Shape           =   5  'Rounded Square
      Top             =   8940
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   10
      Left            =   9240
      Shape           =   5  'Rounded Square
      Top             =   10140
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   9
      Left            =   8580
      Shape           =   5  'Rounded Square
      Top             =   180
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   8
      Left            =   7260
      Shape           =   5  'Rounded Square
      Top             =   2400
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   7
      Left            =   2220
      Shape           =   5  'Rounded Square
      Top             =   4740
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   6
      Left            =   2880
      Shape           =   5  'Rounded Square
      Top             =   420
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   5
      Left            =   10680
      Shape           =   5  'Rounded Square
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   4
      Left            =   6000
      Shape           =   5  'Rounded Square
      Top             =   10140
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   5400
      Shape           =   5  'Rounded Square
      Top             =   900
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   12900
      Shape           =   5  'Rounded Square
      Top             =   3420
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   3780
      Shape           =   5  'Rounded Square
      Top             =   3180
      Width           =   495
   End
   Begin VB.Shape shpRock 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   12840
      Shape           =   5  'Rounded Square
      Top             =   300
      Width           =   495
   End
   Begin VB.Label lblSpeed 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   13800
      TabIndex        =   0
      Top             =   60
      Width           =   1395
   End
   Begin VB.Shape shpCar 
      BorderColor     =   &H000080FF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   1155
      Left            =   7260
      Top             =   5460
      Width           =   615
   End
   Begin VB.Shape shpStartline 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   -60
      Top             =   5340
      Width           =   15270
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "Game"
      Begin VB.Menu mnuRestart 
         Caption         =   "Restart"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "End Game"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Accel As Boolean, Brake As Boolean, Reverse As Boolean
Dim Speed, Speed2, RPM, Gear, ShiftTime, LatMove, Damage
Dim CarRed, CarGreen, FrmGreen, RockRed, RockGreen, RockBlue, Startred, yellow
Dim i, FrontDamage, LeftDamage, RightDamage, RearDamage, Severity, TrailEnd
Dim Outcome

Private Sub cmdMove_KeyPress(KeyAscii As Integer)
    cmdMove.SetFocus
    MsgBox (KeyAscii)
End Sub

Private Sub KeyPress(KeyAscii As Integer)
MsgBox (KeyAscii)
End Sub

Private Sub Form_Load()
TrailEnd = 0
Form1.Height = 11520
Form1.Width = 28040
    shpCar.FillColor = RGB(255, 200, 0)
    shpCar.BorderColor = RGB(255, 200, 0)
    For i = 0 To 49
        shpRock(i).FillColor = RGB(120, 120, 120)
        shpRock(i).BorderColor = RGB(120, 120, 120)
        Next i
        shpStartline.FillColor = RGB(255, 0, 0)
        shpStartline.BorderColor = RGB(255, 0, 0)
        Form1.BackColor = RGB(0, 200, 0)
        For i = 22 To 49
        shpRock(i).Top = shpRock(i).Top + 11520
        Next i
Speed = 0
Gear = 1
RPM = 0
LatMove = 0
CarRed = 255
CarGreen = 100
FrmGreen = 200
RockRed = 120
RockGreen = 120
RockBlue = 120
Startred = 255
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuPause_Click()
If tmrMain.Enabled = True Then
tmrMain.Enabled = False
Else
tmrMain.Enabled = True
End If
End Sub

Private Sub tmrMain_Timer()
 Dim x, s, random
'If RPM > 5000 Then RPM = 5000
With shpSmoke(TrailEnd)
.Top = shpCar.Top + 855
.Left = shpCar.Left + 120
.FillColor = RGB(254, 254, 254)
.ZOrder (0)
.Width = 495
.Height = 495
End With
If TrailEnd < 39 Then
TrailEnd = TrailEnd + 1
Else
TrailEnd = 0
End If
For s = 0 To 39
Select Case shpSmoke(s).FillColor
Case Is = RGB(254, 254, 254)
shpSmoke(s).FillColor = RGB(255, 255, 255)
shpSmoke(s).BorderColor = RGB(255, 255, 255)
shpSmoke(s).Width = 530
Case Is = RGB(255, 255, 255)
shpSmoke(s).FillColor = RGB(255, 255, 100)
shpSmoke(s).BorderColor = RGB(255, 255, 100)
shpSmoke(s).Width = 560
Case Is = RGB(255, 255, 100)
shpSmoke(s).FillColor = RGB(255, 220, 0)
shpSmoke(s).BorderColor = RGB(255, 220, 0)
shpSmoke(s).Width = 590
Case Is = RGB(255, 220, 0)
shpSmoke(s).FillColor = RGB(255, 180, 0)
shpSmoke(s).BorderColor = RGB(255, 180, 0)
shpSmoke(s).Width = 580
shpSmoke(s).Height = 475
Case Is = RGB(255, 180, 0)
shpSmoke(s).FillColor = RGB(255, 90, 0)
shpSmoke(s).BorderColor = RGB(255, 90, 0)
shpSmoke(s).Width = 570
shpSmoke(s).Height = 450
Case Is = RGB(255, 90, 0)
shpSmoke(s).FillColor = RGB(255, 0, 0)
shpSmoke(s).BorderColor = RGB(255, 0, 0)
shpSmoke(s).Width = 560
shpSmoke(s).Height = 425
Case Is = RGB(255, 0, 0)
shpSmoke(s).FillColor = RGB(150, 50, 50)
shpSmoke(s).BorderColor = RGB(150, 50, 50)
shpSmoke(s).Width = 550
shpSmoke(s).Height = 400
Case Is = RGB(150, 50, 50)
shpSmoke(s).FillColor = RGB(100, 90, 90)
shpSmoke(s).BorderColor = RGB(100, 90, 90)
shpSmoke(s).Width = 540
shpSmoke(s).Height = 375
Case Is = RGB(100, 90, 90)
random = Int(Rnd * 90)
shpSmoke(s).FillColor = RGB(random, random, random)
shpSmoke(s).BorderColor = RGB(random, random, random)
shpSmoke(s).Width = 530
shpSmoke(s).Height = 350
Case Else
shpSmoke(s).Width = shpSmoke(s).Width - 10
If shpSmoke(s).Height > 25 Then shpSmoke(s).Height = shpSmoke(s).Height - 25
End Select
shpSmoke(s).Top = shpSmoke(s).Top + Speed
Next s
If Speed > 200 Then Speed = 200
If Speed2 < 120 Then Speed2 = Speed2 + 2
shpPlyr2.Top = shpPlyr2.Top + Speed * 2 - Speed2 * 2
If shpPlyr2.Top < -2000 Then shpPlyr2.Top = 23040
If shpPlyr2.Top > 23040 Then shpPlyr2.Top = -2000
If Accel = True Then
    'RPM = RPM + 45
'Select Case Gear
    'Case Is = 1
'   'Speed = Int(RPM / 60)
'    Case Is = 2
'    Speed = Int(RPM / 45)
'    Case Is = 3
'    Speed = Int(RPM / 30)
'    Case Is = 4
'    Speed = Int(RPM / 20)
'End Select
Speed = Speed + 2
End If
If Reverse = False Then
shpStartline.Top = shpStartline.Top + Speed * 2
If shpStartline.Top >= 23040 Then shpStartline.Top = -2000
Else
shpStartline.Top = shpStartline.Top - Speed * 2
If shpStartline.Top <= -2000 Then shpStartline.Top = 23040
End If
    For i = 0 To 49
    If Reverse = False Then
        shpRock(i).Top = shpRock(i).Top + Speed * 2
        If shpRock(i).Top >= 23040 Then shpRock(i).Top = -2000
        Else
        shpRock(i).Top = shpRock(i).Top - Speed * 2
        If shpRock(i).Top <= -735 Then shpRock(i).Top = 23040
        End If
        If shpCar.Top < shpRock(i).Top + 495 And shpCar.Top > shpRock(i).Top - 1155 Then
        If shpCar.Left < shpRock(i).Left + 495 And shpCar.Left > shpRock(i).Left - 615 Then
        If shpCar.Top >= shpRock(i).Top And shpCar.Top < shpRock(i).Top + 495 Then
        shpRock(i).Top = shpCar.Top - 495
        End If
        If shpCar.Top + 1155 > shpRock(i).Top And shpCar.Top + 1155 <= shpRock(i).Top + 495 Then
        shpRock(i).Top = shpCar.Top + 1155
        End If
        If shpCar.Top < shpRock(i).Top And shpCar.Top + 1155 > shpRock(i).Top + 495 Then
        If shpCar.Left + 615 >= shpRock(i).Left And shpCar.Left + 615 < shpRock(i).Left + 495 Then
        shpRock(i).Left = shpCar.Left + 615
        Else
        shpRock(i).Left = shpCar.Left - 495
        End If
        End If
        Crash
        End If
        End If
        If shpRock(i).Left >= shpPlyr2.Left - 495 And shpRock(i).Left <= shpPlyr2.Left + 615 Then
        If shpRock(i).Top > shpPlyr2.Top - 2505 And shpRock(i).Top < shpPlyr2.Top Then
If shpRock(i).Left + 247 > shpPlyr2.Left + 307 Then 'And shpRock(i).Left > shpPlyr2.Left + 615 Then
shpPlyr2.Left = shpPlyr2.Left - Speed2 * 2
End If
If shpRock(i).Left + 248 < shpPlyr2.Left + 307 And shpRock(i).Left + 495 > shpPlyr2.Left Then
shpPlyr2.Left = shpPlyr2.Left + Speed2 * 2
End If
End If
End If
    Next i
If Brake = True Then
    Speed = Speed - 5
If Speed < 0 Then Speed = 0
End If
    Select Case LatMove
    Case Is = -5
    shpCar.Left = shpCar.Left - Speed * 2
    Case Is = -4
    shpCar.Left = shpCar.Left - Speed * 1.5
    Case Is = -3
    shpCar.Left = shpCar.Left - Speed * 1
    Case Is = -2
    shpCar.Left = shpCar.Left - Speed * 0.5
    Case Is = -1
    shpCar.Left = shpCar.Left - Speed * 0.2
    Case Is = 1
    shpCar.Left = shpCar.Left + Speed * 0.2
    Case Is = 5
    shpCar.Left = shpCar.Left + Speed * 2
    Case Is = 4
    shpCar.Left = shpCar.Left + Speed * 1.5
    Case Is = 3
    shpCar.Left = shpCar.Left + Speed * 1
    Case Is = 2
    shpCar.Left = shpCar.Left + Speed * 0.5
    End Select
lblSpeed.Caption = Speed
txtMove.Text = ""
End Sub

Private Sub tmrPwrDn_Timer()


Startred = Startred - 10
CarRed = CarRed - 10
CarGreen = CarGreen - 10
FrmGreen = FrmGreen - 10
RockRed = RockRed - 10
RockGreen = RockGreen - 10
RockBlue = RockBlue - 10
If CarRed < 0 Then
CarRed = 0
End If
If CarGreen < 0 Then
CarGreen = 0
End If
If FrmGreen < 0 Then
FrmGreen = 0
End If
If RockRed < 0 Then
RockRed = 0
End If
If RockGreen < 0 Then
RockGreen = 0
End If
If RockBlue < 0 Then
RockBlue = 0
End If
If Startred < 0 Then
Startred = 0
End If
shpCar.FillColor = RGB(CarRed, CarGreen, 0)
    shpCar.BorderColor = RGB(CarRed, CarGreen, 0)
    For i = 0 To 49
        shpRock(i).FillColor = RGB(RockRed, RockGreen, RockBlue)
        shpRock(i).BorderColor = RGB(RockRed, RockGreen, RockBlue)
        Next i
        shpStartline.FillColor = RGB(Startred, 0, 0)
        shpStartline.BorderColor = RGB(Startred, 0, 0)
                Form1.BackColor = RGB(0, FrmGreen, 0)
If Startred = 0 Then
yellow = yellow + 10
lblInfo.Caption = "GAME OVER"
lblInfo.ForeColor = RGB(yellow, yellow, 0)
End If
End Sub

Private Sub txtMove_Keypress(KeyAscii As Integer)
    txtMove.SetFocus
    MsgBox (KeyAscii)
End Sub
Private Sub txtMove_Keydown(Keycode As Integer, Shift As Integer)
'Up=38
'Lft=37
'Rt=39
'Dn=40
Select Case Keycode
Case Is = 38
Accel = True
Case Is = 40
Brake = True
Case Is = 37
If LatMove > -4 Then LatMove = LatMove - 1
Case Is = 39
If LatMove < 4 Then LatMove = LatMove + 1
Case Is = 16
'Pressed Shift
Reverse = True
End Select
End Sub
Private Sub txtMove_Keyup(Keycode As Integer, Shift As Integer)
Select Case Keycode
Case Is = 38
Accel = False
Case Is = 40
Brake = False
Case Is = 16
Reverse = False
End Select
End Sub

Private Sub Crash()
Dim CrashAngle
Select Case Speed
Case Is > 150
FrontDamage = FrontDamage + 3
Severity = "destroyed"
Case Is > 100
FrontDamage = FrontDamage + 2
Severity = "crumpled"
Case Is > 50
FrontDamage = FrontDamage + 1
Severity = "dented"
Case Else
Speed = 0
Exit Sub
End Select
Speed = 0
If shpCar.Left + 307 < shpRock(i).Left + 495 And shpCar.Left + 307 > shpRock(i).Left And shpRock(i).Top + 495 < shpCar.Top + 600 Then
CrashAngle = "front bumper"

Else
If shpCar.Left + 307 < shpRock(i).Left Then
CrashAngle = "right side"
Else
If shpCar.Left + 307 > shpRock(i).Left + 495 Then
CrashAngle = "left side"
Else
CrashAngle = "rear bumper"
End If
End If
End If
lblInfo.FontSize = 14

lblInfo.Caption = lblInfo.Caption & vbLf & "You " & Severity & " your " & CrashAngle & " on a rock.  Slow down!"
'        If Speed > 50 And Speed < 100 Then Damage = Damage + 1
'        If Speed >= 100 Then Damage = Damage + 2
'        Speed = 0
'        Select Case Damage
'        Case Is = 1
'        shpDamage(4).FillColor = RGB(0, 0, 0)
'        Case Is = 2
'        shpDamage(4).FillColor = RGB(0, 0, 0)
'        shpDamage(3).FillColor = RGB(0, 0, 0)
'        shpDamage(2).FillColor = RGB(255, 255, 0)
'        shpDamage(1).FillColor = RGB(255, 255, 0)
'        shpDamage(0).FillColor = RGB(255, 255, 0)
'        Case Is = 3
'        shpDamage(4).FillColor = RGB(0, 0, 0)
'        shpDamage(3).FillColor = RGB(0, 0, 0)
'        shpDamage(2).FillColor = RGB(0, 0, 0)
'        shpDamage(1).FillColor = RGB(255, 255, 0)
'        shpDamage(0).FillColor = RGB(255, 255, 0)
'        Case Is = 4
'        shpDamage(4).FillColor = RGB(0, 0, 0)
'        shpDamage(3).FillColor = RGB(0, 0, 0)
'        shpDamage(2).FillColor = RGB(0, 0, 0)
'        shpDamage(1).FillColor = RGB(0, 0, 0)
'        shpDamage(0).FillColor = RGB(255, 0, 0)
'        Case Is >= 5
'        shpDamage(4).FillColor = RGB(0, 0, 0)
'        shpDamage(3).FillColor = RGB(0, 0, 0)
'        shpDamage(2).FillColor = RGB(0, 0, 0)
'        shpDamage(1).FillColor = RGB(0, 0, 0)
'        shpDamage(0).FillColor = RGB(0, 0, 0)
'        tmrPwrDn.Enabled = True
'        tmrMain.Enabled = False
'        Outcome = "GAME OVER"
'        End Select
End Sub
