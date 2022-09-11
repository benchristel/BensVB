VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00004000&
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBorders 
      Caption         =   "Turn Borders Off"
      Height          =   375
      Left            =   9120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   14
      Left            =   8280
      Picture         =   "frmMain.frx":0000
      Top             =   3840
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   13
      Left            =   8280
      Picture         =   "frmMain.frx":0101
      Top             =   2400
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   12
      Left            =   10440
      Picture         =   "frmMain.frx":01FD
      Top             =   1680
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   11
      Left            =   9000
      Picture         =   "frmMain.frx":02F5
      Top             =   1680
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   10
      Left            =   9720
      Picture         =   "frmMain.frx":03EF
      Top             =   1680
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   9
      Left            =   8280
      Picture         =   "frmMain.frx":04D6
      Top             =   3120
      Width           =   720
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FF80FF&
      Height          =   375
      Index           =   39
      Left            =   10560
      TabIndex        =   40
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FF00FF&
      Height          =   375
      Index           =   38
      Left            =   10200
      TabIndex        =   39
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00C000C0&
      Height          =   375
      Index           =   37
      Left            =   9840
      TabIndex        =   38
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00800080&
      Height          =   375
      Index           =   36
      Left            =   9480
      TabIndex        =   37
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00400040&
      Height          =   375
      Index           =   35
      Left            =   9120
      TabIndex        =   36
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FF8080&
      Height          =   375
      Index           =   34
      Left            =   10560
      TabIndex        =   35
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FF0000&
      Height          =   375
      Index           =   33
      Left            =   10200
      TabIndex        =   34
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00C00000&
      Height          =   375
      Index           =   32
      Left            =   9840
      TabIndex        =   33
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00800000&
      Height          =   375
      Index           =   31
      Left            =   9480
      TabIndex        =   32
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00400000&
      Height          =   375
      Index           =   30
      Left            =   9120
      TabIndex        =   31
      Top             =   6720
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FFFF80&
      Height          =   375
      Index           =   29
      Left            =   10560
      TabIndex        =   30
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FFFF00&
      Height          =   375
      Index           =   28
      Left            =   10200
      TabIndex        =   29
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00C0C000&
      Height          =   375
      Index           =   27
      Left            =   9840
      TabIndex        =   28
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00808000&
      Height          =   375
      Index           =   26
      Left            =   9480
      TabIndex        =   27
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00404000&
      Height          =   375
      Index           =   25
      Left            =   9120
      TabIndex        =   26
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H0080FF80&
      Height          =   375
      Index           =   9
      Left            =   10560
      TabIndex        =   10
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H0000FF00&
      Height          =   375
      Index           =   8
      Left            =   10200
      TabIndex        =   9
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H0000C000&
      Height          =   375
      Index           =   7
      Left            =   9840
      TabIndex        =   8
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00008000&
      Height          =   375
      Index           =   6
      Left            =   9480
      TabIndex        =   7
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00004000&
      Height          =   375
      Index           =   5
      Left            =   9120
      TabIndex        =   6
      Top             =   6000
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Index           =   22
      Left            =   10560
      TabIndex        =   23
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H0000C0C0&
      Height          =   375
      Index           =   21
      Left            =   9840
      TabIndex        =   22
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00008080&
      Height          =   375
      Index           =   20
      Left            =   9480
      TabIndex        =   21
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00004040&
      Height          =   375
      Index           =   19
      Left            =   9120
      TabIndex        =   20
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H0000FFFF&
      Height          =   375
      Index           =   4
      Left            =   10200
      TabIndex        =   5
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H0080C0FF&
      Height          =   375
      Index           =   23
      Left            =   10560
      TabIndex        =   24
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H000040C0&
      Height          =   375
      Index           =   18
      Left            =   9840
      TabIndex        =   19
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00004080&
      Height          =   375
      Index           =   17
      Left            =   9480
      TabIndex        =   18
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00404080&
      Height          =   375
      Index           =   16
      Left            =   9120
      TabIndex        =   17
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H000080FF&
      Height          =   375
      Index           =   3
      Left            =   10200
      TabIndex        =   4
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H008080FF&
      Height          =   375
      Index           =   24
      Left            =   10560
      TabIndex        =   25
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H000000C0&
      Height          =   375
      Index           =   15
      Left            =   9840
      TabIndex        =   16
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00000080&
      Height          =   375
      Index           =   14
      Left            =   9480
      TabIndex        =   15
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00000040&
      Height          =   375
      Index           =   13
      Left            =   9120
      TabIndex        =   14
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   10200
      TabIndex        =   3
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   12
      Left            =   10560
      TabIndex        =   13
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00808080&
      Height          =   375
      Index           =   11
      Left            =   9840
      TabIndex        =   12
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00404040&
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   11
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   9120
      TabIndex        =   2
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   0
      Left            =   10200
      TabIndex        =   1
      Top             =   4560
      Width           =   375
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   8
      Left            =   10410
      Picture         =   "frmMain.frx":05D8
      Top             =   3810
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   7
      Left            =   9705
      Picture         =   "frmMain.frx":06E1
      Top             =   3810
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   6
      Left            =   9000
      Picture         =   "frmMain.frx":07D5
      Top             =   3810
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   5
      Left            =   10410
      Picture         =   "frmMain.frx":08D6
      Top             =   3105
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   4
      Left            =   9705
      Picture         =   "frmMain.frx":09DB
      Top             =   3105
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   3
      Left            =   9000
      Picture         =   "frmMain.frx":0AD0
      Top             =   3105
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   2
      Left            =   10410
      Picture         =   "frmMain.frx":0BD1
      Top             =   2400
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   1
      Left            =   9705
      Picture         =   "frmMain.frx":0CCF
      Top             =   2400
      Width           =   720
   End
   Begin VB.Image imgBrush 
      Height          =   720
      Index           =   0
      Left            =   9000
      Picture         =   "frmMain.frx":0DBC
      Top             =   2400
      Width           =   720
   End
   Begin VB.Image imgKnot 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   720
      Index           =   0
      Left            =   240
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BrushIndex As Integer, BordersOn As Boolean

Private Sub cmdBorders_Click()
Dim i
If BordersOn = True Then
BordersOn = False
For i = 1 To 100
imgKnot(i).BorderStyle = 0
Next i
cmdBorders.Caption = "Turn Borders On"
Else
BordersOn = True
For i = 1 To 100
imgKnot(i).BorderStyle = 1
Next i
cmdBorders.Caption = "Turn Borders Off"
End If

End Sub

Private Sub Form_Load()
Dim i, j, Index As Integer
For i = 1 To 10
For j = 1 To 10
Index = Index + 1
Load imgKnot(Index)
imgKnot(Index).Left = i * 47
imgKnot(Index).Top = j * 47
imgKnot(Index).Height = 50
imgKnot(Index).Width = 50
imgKnot(Index).Visible = True
Next j
Next i
BordersOn = True
End Sub

Private Sub imgBrush_Click(Index As Integer)
Dim i
BrushIndex = Index
For i = 0 To 14
imgBrush(i).BorderStyle = 0
Next i
imgBrush(Index).BorderStyle = 1
End Sub

Private Sub imgKnot_Click(Index As Integer)
imgKnot(Index).Picture = imgBrush(BrushIndex).Picture
End Sub


Private Sub imgKnot_DblClick(Index As Integer)
imgKnot(Index).Picture = imgKnot(0).Picture
End Sub

Private Sub lblColor_Click(Index As Integer)
frmMain.BackColor = lblColor(Index).BackColor
End Sub
