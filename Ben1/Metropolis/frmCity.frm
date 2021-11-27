VERSION 5.00
Begin VB.Form frmCity 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   1800
      Top             =   60
   End
   Begin VB.Image imgMove 
      Height          =   480
      Left            =   14520
      Picture         =   "frmCity.frx":0000
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgBuild 
      Height          =   480
      Left            =   14520
      Picture         =   "frmCity.frx":030A
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgTarget 
      Height          =   960
      Index           =   0
      Left            =   960
      Picture         =   "frmCity.frx":12CC
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgCitizen 
      Height          =   960
      Index           =   0
      Left            =   60
      Picture         =   "frmCity.frx":15D6
      Stretch         =   -1  'True
      Top             =   60
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   164
      Left            =   13500
      Picture         =   "frmCity.frx":18E0
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   163
      Left            =   12540
      Picture         =   "frmCity.frx":1BEA
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   162
      Left            =   11580
      Picture         =   "frmCity.frx":1EF4
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   161
      Left            =   10620
      Picture         =   "frmCity.frx":21FE
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   160
      Left            =   9660
      Picture         =   "frmCity.frx":2508
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   159
      Left            =   8700
      Picture         =   "frmCity.frx":2812
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   158
      Left            =   7740
      Picture         =   "frmCity.frx":2B1C
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   157
      Left            =   6780
      Picture         =   "frmCity.frx":2E26
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   156
      Left            =   5820
      Picture         =   "frmCity.frx":3130
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   155
      Left            =   4860
      Picture         =   "frmCity.frx":343A
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   154
      Left            =   3900
      Picture         =   "frmCity.frx":3744
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   153
      Left            =   2940
      Picture         =   "frmCity.frx":3A4E
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   152
      Left            =   1980
      Picture         =   "frmCity.frx":3D58
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   151
      Left            =   1020
      Picture         =   "frmCity.frx":4062
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   150
      Left            =   60
      Picture         =   "frmCity.frx":436C
      Stretch         =   -1  'True
      Top             =   9660
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   149
      Left            =   13500
      Picture         =   "frmCity.frx":4676
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   148
      Left            =   12540
      Picture         =   "frmCity.frx":4980
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   147
      Left            =   11580
      Picture         =   "frmCity.frx":4C8A
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   146
      Left            =   10620
      Picture         =   "frmCity.frx":4F94
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   145
      Left            =   9660
      Picture         =   "frmCity.frx":529E
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   144
      Left            =   8700
      Picture         =   "frmCity.frx":55A8
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   143
      Left            =   7740
      Picture         =   "frmCity.frx":58B2
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   142
      Left            =   6780
      Picture         =   "frmCity.frx":5BBC
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   141
      Left            =   5820
      Picture         =   "frmCity.frx":5EC6
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   140
      Left            =   4860
      Picture         =   "frmCity.frx":61D0
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   139
      Left            =   3900
      Picture         =   "frmCity.frx":64DA
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   138
      Left            =   2940
      Picture         =   "frmCity.frx":67E4
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   137
      Left            =   1980
      Picture         =   "frmCity.frx":6AEE
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   136
      Left            =   1020
      Picture         =   "frmCity.frx":6DF8
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   135
      Left            =   60
      Picture         =   "frmCity.frx":7102
      Stretch         =   -1  'True
      Top             =   8700
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   134
      Left            =   13500
      Picture         =   "frmCity.frx":740C
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   133
      Left            =   12540
      Picture         =   "frmCity.frx":7716
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   132
      Left            =   11580
      Picture         =   "frmCity.frx":7A20
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   131
      Left            =   10620
      Picture         =   "frmCity.frx":7D2A
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   130
      Left            =   9660
      Picture         =   "frmCity.frx":8034
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   129
      Left            =   8700
      Picture         =   "frmCity.frx":833E
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   128
      Left            =   7740
      Picture         =   "frmCity.frx":8648
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   127
      Left            =   6780
      Picture         =   "frmCity.frx":8952
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   126
      Left            =   5820
      Picture         =   "frmCity.frx":8C5C
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   125
      Left            =   4860
      Picture         =   "frmCity.frx":8F66
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   124
      Left            =   3900
      Picture         =   "frmCity.frx":9270
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   123
      Left            =   2940
      Picture         =   "frmCity.frx":957A
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   122
      Left            =   1980
      Picture         =   "frmCity.frx":9884
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   121
      Left            =   1020
      Picture         =   "frmCity.frx":9B8E
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   120
      Left            =   60
      Picture         =   "frmCity.frx":9E98
      Stretch         =   -1  'True
      Top             =   7740
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   119
      Left            =   13500
      Picture         =   "frmCity.frx":A1A2
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   118
      Left            =   12540
      Picture         =   "frmCity.frx":A4AC
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   117
      Left            =   11580
      Picture         =   "frmCity.frx":A7B6
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   116
      Left            =   10620
      Picture         =   "frmCity.frx":AAC0
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   115
      Left            =   9660
      Picture         =   "frmCity.frx":ADCA
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   114
      Left            =   8700
      Picture         =   "frmCity.frx":B0D4
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   113
      Left            =   7740
      Picture         =   "frmCity.frx":B3DE
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   112
      Left            =   6780
      Picture         =   "frmCity.frx":B6E8
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   111
      Left            =   5820
      Picture         =   "frmCity.frx":B9F2
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   110
      Left            =   4860
      Picture         =   "frmCity.frx":BCFC
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   109
      Left            =   3900
      Picture         =   "frmCity.frx":C006
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   108
      Left            =   2940
      Picture         =   "frmCity.frx":C310
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   107
      Left            =   1980
      Picture         =   "frmCity.frx":C61A
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   106
      Left            =   1020
      Picture         =   "frmCity.frx":C924
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   105
      Left            =   60
      Picture         =   "frmCity.frx":CC2E
      Stretch         =   -1  'True
      Top             =   6780
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   104
      Left            =   13500
      Picture         =   "frmCity.frx":CF38
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   103
      Left            =   12540
      Picture         =   "frmCity.frx":D242
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   102
      Left            =   11580
      Picture         =   "frmCity.frx":D54C
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   101
      Left            =   10620
      Picture         =   "frmCity.frx":D856
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   100
      Left            =   9660
      Picture         =   "frmCity.frx":DB60
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   99
      Left            =   8700
      Picture         =   "frmCity.frx":DE6A
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   98
      Left            =   7740
      Picture         =   "frmCity.frx":E174
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   97
      Left            =   6780
      Picture         =   "frmCity.frx":E47E
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   96
      Left            =   5820
      Picture         =   "frmCity.frx":E788
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   95
      Left            =   4860
      Picture         =   "frmCity.frx":EA92
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   94
      Left            =   3900
      Picture         =   "frmCity.frx":ED9C
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   93
      Left            =   2940
      Picture         =   "frmCity.frx":F0A6
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   92
      Left            =   1980
      Picture         =   "frmCity.frx":F3B0
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   91
      Left            =   1020
      Picture         =   "frmCity.frx":F6BA
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   90
      Left            =   60
      Picture         =   "frmCity.frx":F9C4
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   89
      Left            =   13500
      Picture         =   "frmCity.frx":FCCE
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   88
      Left            =   12540
      Picture         =   "frmCity.frx":FFD8
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   87
      Left            =   11580
      Picture         =   "frmCity.frx":102E2
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   86
      Left            =   10620
      Picture         =   "frmCity.frx":105EC
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   85
      Left            =   9660
      Picture         =   "frmCity.frx":108F6
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   84
      Left            =   8700
      Picture         =   "frmCity.frx":10C00
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   83
      Left            =   7740
      Picture         =   "frmCity.frx":10F0A
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   82
      Left            =   6780
      Picture         =   "frmCity.frx":11214
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   81
      Left            =   5820
      Picture         =   "frmCity.frx":1151E
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   80
      Left            =   4860
      Picture         =   "frmCity.frx":11828
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   79
      Left            =   3900
      Picture         =   "frmCity.frx":11B32
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   78
      Left            =   2940
      Picture         =   "frmCity.frx":11E3C
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   77
      Left            =   1980
      Picture         =   "frmCity.frx":12146
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   76
      Left            =   1020
      Picture         =   "frmCity.frx":12450
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   75
      Left            =   60
      Picture         =   "frmCity.frx":1275A
      Stretch         =   -1  'True
      Top             =   4860
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   74
      Left            =   13500
      Picture         =   "frmCity.frx":12A64
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   73
      Left            =   12540
      Picture         =   "frmCity.frx":12D6E
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   72
      Left            =   11580
      Picture         =   "frmCity.frx":13078
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   71
      Left            =   10620
      Picture         =   "frmCity.frx":13382
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   70
      Left            =   9660
      Picture         =   "frmCity.frx":1368C
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   69
      Left            =   8700
      Picture         =   "frmCity.frx":13996
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   68
      Left            =   7740
      Picture         =   "frmCity.frx":13CA0
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   67
      Left            =   6780
      Picture         =   "frmCity.frx":13FAA
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   66
      Left            =   5820
      Picture         =   "frmCity.frx":142B4
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   65
      Left            =   4860
      Picture         =   "frmCity.frx":145BE
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   64
      Left            =   3900
      Picture         =   "frmCity.frx":148C8
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   63
      Left            =   2940
      Picture         =   "frmCity.frx":14BD2
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   62
      Left            =   1980
      Picture         =   "frmCity.frx":14EDC
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   61
      Left            =   1020
      Picture         =   "frmCity.frx":151E6
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   60
      Left            =   60
      Picture         =   "frmCity.frx":154F0
      Stretch         =   -1  'True
      Top             =   3900
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   59
      Left            =   13500
      Picture         =   "frmCity.frx":157FA
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   58
      Left            =   12540
      Picture         =   "frmCity.frx":15B04
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   57
      Left            =   11580
      Picture         =   "frmCity.frx":15E0E
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   56
      Left            =   10620
      Picture         =   "frmCity.frx":16118
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   55
      Left            =   9660
      Picture         =   "frmCity.frx":16422
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   54
      Left            =   8700
      Picture         =   "frmCity.frx":1672C
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   53
      Left            =   7740
      Picture         =   "frmCity.frx":16A36
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   52
      Left            =   6780
      Picture         =   "frmCity.frx":16D40
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   51
      Left            =   5820
      Picture         =   "frmCity.frx":1704A
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   50
      Left            =   4860
      Picture         =   "frmCity.frx":17354
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   49
      Left            =   3900
      Picture         =   "frmCity.frx":1765E
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   48
      Left            =   2940
      Picture         =   "frmCity.frx":17968
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   47
      Left            =   1980
      Picture         =   "frmCity.frx":17C72
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   46
      Left            =   1020
      Picture         =   "frmCity.frx":17F7C
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   45
      Left            =   60
      Picture         =   "frmCity.frx":18286
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   44
      Left            =   13500
      Picture         =   "frmCity.frx":18590
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   43
      Left            =   12540
      Picture         =   "frmCity.frx":1889A
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   42
      Left            =   11580
      Picture         =   "frmCity.frx":18BA4
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   41
      Left            =   10620
      Picture         =   "frmCity.frx":18EAE
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   40
      Left            =   9660
      Picture         =   "frmCity.frx":191B8
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   39
      Left            =   8700
      Picture         =   "frmCity.frx":194C2
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   38
      Left            =   7740
      Picture         =   "frmCity.frx":197CC
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   37
      Left            =   6780
      Picture         =   "frmCity.frx":19AD6
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   36
      Left            =   5820
      Picture         =   "frmCity.frx":19DE0
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   35
      Left            =   4860
      Picture         =   "frmCity.frx":1A0EA
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   34
      Left            =   3900
      Picture         =   "frmCity.frx":1A3F4
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   33
      Left            =   2940
      Picture         =   "frmCity.frx":1A6FE
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   32
      Left            =   1980
      Picture         =   "frmCity.frx":1AA08
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   31
      Left            =   1020
      Picture         =   "frmCity.frx":1AD12
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   30
      Left            =   60
      Picture         =   "frmCity.frx":1B01C
      Stretch         =   -1  'True
      Top             =   1980
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   29
      Left            =   13500
      Picture         =   "frmCity.frx":1B326
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   28
      Left            =   12540
      Picture         =   "frmCity.frx":1B630
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   27
      Left            =   11580
      Picture         =   "frmCity.frx":1B93A
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   26
      Left            =   10620
      Picture         =   "frmCity.frx":1BC44
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   25
      Left            =   9660
      Picture         =   "frmCity.frx":1BF4E
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   24
      Left            =   8700
      Picture         =   "frmCity.frx":1C258
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   23
      Left            =   7740
      Picture         =   "frmCity.frx":1C562
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   22
      Left            =   6780
      Picture         =   "frmCity.frx":1C86C
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   21
      Left            =   5820
      Picture         =   "frmCity.frx":1CB76
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   20
      Left            =   4860
      Picture         =   "frmCity.frx":1CE80
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   19
      Left            =   3900
      Picture         =   "frmCity.frx":1D18A
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   18
      Left            =   2940
      Picture         =   "frmCity.frx":1D494
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   17
      Left            =   1980
      Picture         =   "frmCity.frx":1D79E
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   16
      Left            =   1020
      Picture         =   "frmCity.frx":1DAA8
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   15
      Left            =   60
      Picture         =   "frmCity.frx":1DDB2
      Stretch         =   -1  'True
      Top             =   1020
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   14
      Left            =   13500
      Picture         =   "frmCity.frx":1E0BC
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   13
      Left            =   12540
      Picture         =   "frmCity.frx":1E3C6
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   12
      Left            =   11580
      Picture         =   "frmCity.frx":1E6D0
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   11
      Left            =   10620
      Picture         =   "frmCity.frx":1E9DA
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   10
      Left            =   9660
      Picture         =   "frmCity.frx":1ECE4
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   9
      Left            =   8700
      Picture         =   "frmCity.frx":1EFEE
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   8
      Left            =   7740
      Picture         =   "frmCity.frx":1F2F8
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   7
      Left            =   6780
      Picture         =   "frmCity.frx":1F602
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   6
      Left            =   5820
      Picture         =   "frmCity.frx":1F90C
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   5
      Left            =   4860
      Picture         =   "frmCity.frx":1FC16
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   4
      Left            =   3900
      Picture         =   "frmCity.frx":1FF20
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   3
      Left            =   2940
      Picture         =   "frmCity.frx":2022A
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   2
      Left            =   1980
      Picture         =   "frmCity.frx":20534
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   1
      Left            =   1020
      Picture         =   "frmCity.frx":2083E
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
   Begin VB.Image imgTile 
      Height          =   960
      Index           =   0
      Left            =   60
      Picture         =   "frmCity.frx":20B48
      Stretch         =   -1  'True
      Top             =   60
      Width           =   960
   End
End
Attribute VB_Name = "frmCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Citizens, Active, Tool

Private Sub Form_Load()
Dim Startspace
Randomize
Citizens = 0
Citizens = Citizens + 1
Load imgCitizen(Citizens)
Load imgTarget(Citizens)
imgCitizen(Citizens).Visible = True
imgCitizen(Citizens).ZOrder
Startspace = Int(Rnd * 164)
imgCitizen(Citizens).Top = imgTile(Startspace).Top
imgCitizen(Citizens).Left = imgTile(Startspace).Left
imgTarget(Citizens).Top = imgTile(Startspace).Top
imgTarget(Citizens).Left = imgTile(Startspace).Left
End Sub

Private Sub imgCitizen_DblClick(Index As Integer)
Active = Index
End Sub

Private Sub imgMove_Click()
Tool = "Selector"
frmCity.MousePointer = 99 'custom pointer
frmCity.MouseIcon = LoadPicture("C:\My Documents\Ben\BensVB\Ben1\Metropolis\Citizen.ico")
End Sub

Private Sub imgTile_Click(Index As Integer)
Select Case Tool
Case Is = "House"
imgTile(Index).Picture = frmToolbox.imgHouse.Picture
Case Is = "Target"
imgTarget(Active).Left = imgTile(Index).Left
imgTarget(Active).Top = imgTile(Index).Top
End If
End Sub

Private Sub tmrTime_Timer()
Dim i
For i = 1 To Citizens
If imgCitizen(i).Top = imgTarget(i).Top Then
    If imgCitizen(i).Left > imgTarget(i).Left Then _
    imgCitizen(i).Left = imgCitizen(i).Left - 960
    If imgCitizen(i).Left < imgTarget(i).Left Then _
    imgCitizen(i).Left = imgCitizen(i).Left + 960
End If
If imgCitizen(i).Top > imgTarget(i).Top Then
    imgCitizen(i).Top = imgCitizen(i).Top - 960
    If imgCitizen(i).Left > imgTarget(i).Left Then _
    imgCitizen(i).Left = imgCitizen(i).Left - 960
    If imgCitizen(i).Left < imgTarget(i).Left Then _
    imgCitizen(i).Left = imgCitizen(i).Left + 960
    End If
If imgCitizen(i).Top < imgTarget(i).Top Then
    imgCitizen(i).Top = imgCitizen(i).Top + 960
    If imgCitizen(i).Left > imgTarget(i).Left Then _
    imgCitizen(i).Left = imgCitizen(i).Left - 960
    If imgCitizen(i).Left < imgTarget(i).Left Then _
    imgCitizen(i).Left = imgCitizen(i).Left + 960
    End If
    Next i
End Sub
