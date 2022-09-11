VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNeutral 
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   10080
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2760
      Top             =   9060
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   10440
      Width           =   375
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   10080
      Width           =   375
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   10080
      Width           =   375
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   9720
      Width           =   375
   End
   Begin VB.TextBox txtDummy 
      BackColor       =   &H00FFFF00&
      Height          =   285
      Left            =   10080
      TabIndex        =   5
      Top             =   1000
      Width           =   150
   End
   Begin VB.Shape shpPlane 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   7020
      Top             =   5940
      Width           =   1215
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   99
      Left            =   4500
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   98
      Left            =   5760
      Top             =   6000
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   97
      Left            =   4680
      Top             =   5940
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   96
      Left            =   4020
      Top             =   5940
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   95
      Left            =   5220
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   94
      Left            =   5460
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   93
      Left            =   5580
      Top             =   6060
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   92
      Left            =   3540
      Top             =   5820
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   91
      Left            =   6000
      Top             =   6060
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   90
      Left            =   6360
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   89
      Left            =   3300
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   88
      Left            =   6540
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   87
      Left            =   4200
      Top             =   5940
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   86
      Left            =   4620
      Top             =   5820
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   85
      Left            =   8640
      Top             =   5940
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   84
      Left            =   6720
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   83
      Left            =   6900
      Top             =   5940
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   82
      Left            =   6600
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   81
      Left            =   6420
      Top             =   5940
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   80
      Left            =   2760
      Top             =   5940
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   79
      Left            =   3840
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   78
      Left            =   4200
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   77
      Left            =   3360
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   76
      Left            =   3180
      Top             =   6060
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   75
      Left            =   2580
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   74
      Left            =   2280
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   73
      Left            =   2340
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   72
      Left            =   5160
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   71
      Left            =   4860
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   70
      Left            =   4020
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   69
      Left            =   5400
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   68
      Left            =   3600
      Top             =   6060
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   67
      Left            =   7440
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   66
      Left            =   3000
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   65
      Left            =   9180
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   64
      Left            =   3000
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   63
      Left            =   2460
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   62
      Left            =   4380
      Top             =   6420
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   61
      Left            =   5040
      Top             =   6360
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   60
      Left            =   7740
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   59
      Left            =   10200
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   58
      Left            =   4680
      Top             =   6420
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   57
      Left            =   5280
      Top             =   6420
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   56
      Left            =   5580
      Top             =   6300
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   55
      Left            =   5820
      Top             =   6360
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   54
      Left            =   6060
      Top             =   6360
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   53
      Left            =   9840
      Top             =   6060
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   52
      Left            =   4860
      Top             =   6480
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   51
      Left            =   5520
      Top             =   6540
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   50
      Left            =   10020
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   49
      Left            =   9420
      Top             =   6000
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   48
      Left            =   9600
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   47
      Left            =   4680
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   46
      Left            =   8280
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   45
      Left            =   6960
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   44
      Left            =   9660
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   43
      Left            =   10500
      Top             =   6360
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   42
      Left            =   10380
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   41
      Left            =   9060
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   40
      Left            =   10560
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   39
      Left            =   10080
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   38
      Left            =   11760
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   37
      Left            =   11640
      Top             =   6420
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   36
      Left            =   11580
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   35
      Left            =   11220
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   34
      Left            =   12000
      Top             =   6240
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   33
      Left            =   7260
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   32
      Left            =   10920
      Top             =   6300
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   31
      Left            =   11400
      Top             =   6480
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   30
      Left            =   11940
      Top             =   6000
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   29
      Left            =   11040
      Top             =   5940
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   28
      Left            =   10680
      Top             =   6300
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   27
      Left            =   11160
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   26
      Left            =   10320
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   25
      Left            =   11520
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   24
      Left            =   11220
      Top             =   6420
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   23
      Left            =   12000
      Top             =   6480
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   22
      Left            =   11760
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   21
      Left            =   11820
      Top             =   6420
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   20
      Left            =   7500
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   19
      Left            =   8460
      Top             =   6180
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   18
      Left            =   8700
      Top             =   6240
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   17
      Left            =   10200
      Top             =   6420
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   16
      Left            =   9840
      Top             =   6300
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   15
      Left            =   10800
      Top             =   6060
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   14
      Left            =   10500
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   13
      Left            =   11040
      Top             =   6540
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   12
      Left            =   12120
      Top             =   6000
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   11
      Left            =   12420
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   10
      Left            =   11340
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   9
      Left            =   7860
      Top             =   6300
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   8
      Left            =   8880
      Top             =   6120
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   7
      Left            =   9240
      Top             =   6240
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   6
      Left            =   8820
      Top             =   5820
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   5
      Left            =   6240
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   4
      Left            =   3720
      Top             =   5880
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   3
      Left            =   5940
      Top             =   5820
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   2
      Left            =   4980
      Top             =   5940
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   0
      Left            =   4380
      Top             =   5820
      Width           =   135
   End
   Begin VB.Shape shpBldg 
      BorderColor     =   &H80000004&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   1
      Left            =   8520
      Top             =   5820
      Width           =   135
   End
   Begin VB.Shape shpControlPanel 
      BorderColor     =   &H00808080&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   60
      Top             =   9540
      Width           =   15135
   End
   Begin VB.Shape shpGround 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   5940
      Left            =   60
      Top             =   6000
      Width           =   15135
   End
   Begin VB.Shape shpSky 
      BorderColor     =   &H00FFFF80&
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   5940
      Left            =   60
      Top             =   60
      Width           =   15135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pitch, VMove(0 To 99), HMove(0 To 99), VMoved(0 To 99), HMoved(0 To 99), MoveDiv

Private Sub txtDummy_Keydown(KeyCode As Integer, Shift As Integer)
MsgBox (KeyCode)
End Sub

Private Sub cmdup_Click()
If Pitch < 5 Then
Pitch = Pitch + 1
End If
End Sub

Private Sub cmdDown_Click()
If Pitch > -5 Then
Pitch = Pitch - 1
End If
End Sub

Private Sub cmdDown_KeyPress(KeyAscii As Integer)
MsgBox ("Keycode is: " & KeyAscii)
End Sub

Private Sub Form_Load()
Dim i, a
For i = 0 To 99
HMove(i) = 7815 - shpBldg(i).Left
VMove(i) = 7815 - HMove(i)
If VMove(i) > HMove(i) Then
HMove(i) = Int(CLng(HMove(i) / VMove(i)))
Else
VMove(i) = Int(CLng(HMove(i) / HMove(i)))
End If
Next i
'    For i = 0 To 99
'    VMove(i) = shpBldg(i).Left - 7815
'    If VMove(i) > 0 Then
'    HMove(i) = 7815 - VMove(i)
'    Else
'    HMove(i) = 7815 + VMove(i)
'    End If
'        If VMove(i) > HMove(i) Then
'        For a = 1 To VMove(i)
'            'If VMove(i) Mod a = 0 And HMove(i) Mod VMove(i) = 0 Then
'            VMove(i) = CLng(Int(VMove(i) / a))
'            HMove(i) = CLng(Int(HMove(i) / VMove(i)))
'            'End If
'        Next a
'        Else
'        For a = 1 To VMove(i)
'            'If HMove(i) Mod a = 0 And VMove(i) Mod HMove(i) = 0 Then
'            HMove(i) = Int(HMove(i)) / a
'            VMove(i) = Int(VMove(i)) / VMove(i)
'            'End If
'        Next a
'        End If
'    Next i
End Sub

Private Sub Timer1_Timer()
Dim i
Select Case Pitch
Case Is = "1"
shpPlane.Top = shpPlane.Top - 2
Case Is = "2"
shpPlane.Top = shpPlane.Top - 4
Case Is = "3"
shpPlane.Top = shpPlane.Top - 6
Case Is = "4"
shpPlane.Top = shpPlane.Top - 8
Case Is = "5"
shpPlane.Top = shpPlane.Top - 10
Case Is = "-1"
shpPlane.Top = shpPlane.Top + 2
Case Is = "-2"
shpPlane.Top = shpPlane.Top + 4
Case Is = "-3"
shpPlane.Top = shpPlane.Top + 6
Case Is = "-4"
shpPlane.Top = shpPlane.Top + 8
Case Is = "-5"
shpPlane.Top = shpPlane.Top + 10
End Select
For i = 0 To 99
If HMoved(i) > HMove(i) Then
VMoved(i) = VMoved(i) + 1
shpBldg(i).Height = shpBldg(i).Height + 1
shpBldg(i).Width = shpBldg(i).Width + 1
shpBldg(i).Top = shpBldg(i).Top - 1
Else
VMoved(i) = 0
End If
If VMoved(i) > VMove(i) Then
HMoved(i) = HMoved(i) + 1
shpBldg(i).Left = shpBldg(i).Left - 1
Else
HMoved(i) = 0
End If
DoEvents
txtDummy.SetFocus
Next i
End Sub
