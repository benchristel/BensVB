VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   Caption         =   "Ben's Squiggle"
   ClientHeight    =   5625
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11850
   BeginProperty Font 
      Name            =   "Papyrus"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Thick Line"
      Height          =   375
      Left            =   10320
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Thin Line"
      Height          =   255
      Left            =   10320
      TabIndex        =   12
      Top             =   3840
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   1
      Top             =   4800
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      ScaleHeight     =   4515
      ScaleWidth      =   10155
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11040
      TabIndex        =   11
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   10
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   9
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11040
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Eraser"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11040
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11040
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuclear 
         Caption         =   "&Clear"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu LineColorsmnu 
      Caption         =   "Line &Colors"
      Begin VB.Menu blackmnu 
         Caption         =   "B&lack"
         Shortcut        =   ^L
      End
      Begin VB.Menu redmnu 
         Caption         =   "&Red"
         Shortcut        =   ^R
      End
      Begin VB.Menu greenmnu 
         Caption         =   "&Green"
         Shortcut        =   ^G
      End
      Begin VB.Menu bluemnu 
         Caption         =   "&Blue"
         Shortcut        =   ^B
      End
      Begin VB.Menu whitemnu 
         Caption         =   "&Eraser"
         Shortcut        =   ^E
      End
      Begin VB.Menu LightBluemnu 
         Caption         =   "L&ight Blue"
         Shortcut        =   ^I
      End
      Begin VB.Menu Yellowmnu 
         Caption         =   "&Yellow"
         Shortcut        =   ^Y
      End
      Begin VB.Menu DarkGreenmnu 
         Caption         =   "&Dark Green"
         Shortcut        =   ^D
      End
      Begin VB.Menu Pinkmnu 
         Caption         =   "&Pink"
         Shortcut        =   ^P
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Drawcol

Private Sub blackmnu_Click()
Drawcol = 0
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 1
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
greenmnu.Checked = False
blackmnu.Checked = True
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = False
Yellowmnu.Checked = False
LightBluemnu.Checked = False
DarkGreenmnu.Checked = False
Pinkmnu.Checked = False
End Sub

Private Sub bluemnu_Click()
Drawcol = 9
Label5.BorderStyle = 0
Label4.BorderStyle = 1
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
greenmnu.Checked = False
blackmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = True
whitemnu.Checked = False
Yellowmnu.Checked = False
LightBluemnu.Checked = False
DarkGreenmnu.Checked = False
Pinkmnu.Checked = False
End Sub

Private Sub Command1_Click()
Picture1.Cls
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub DarkGreenmnu_Click()
Drawcol = 2
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 1
Label9.BorderStyle = 0
blackmnu.Checked = False
greenmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = False
LightBluemnu.Checked = False
Yellowmnu.Checked = False
DarkGreenmnu.Checked = True
Pinkmnu.Checked = False

End Sub

Private Sub Form_Load()
Form1.Visible = False
Form2.Visible = True
Drawcol = 0
blackmnu.Checked = True
Label1.BorderStyle = 1
End Sub

Private Sub greenmnu_Click()
Drawcol = 10
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 1
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
greenmnu.Checked = True
blackmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = False
Yellowmnu.Checked = False
LightBluemnu.Checked = False
DarkGreenmnu.Checked = False
Pinkmnu.Checked = False
End Sub

Private Sub Label1_Click()
Drawcol = 0
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 1
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
blackmnu.Checked = True
greenmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = False
Yellowmnu.Checked = False
LightBluemnu.Checked = False
DarkGreenmnu.Checked = False
Pinkmnu.Checked = False
End Sub

Private Sub Label2_Click()
Drawcol = 12
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 1
Label1.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
blackmnu.Checked = False
greenmnu.Checked = False
redmnu.Checked = True
bluemnu.Checked = False
whitemnu.Checked = False
LightBluemnu.Checked = False
DarkGreenmnu.Checked = False
Yellowmnu.Checked = False
Pinkmnu.Checked = False



End Sub

Private Sub Label3_Click()
Drawcol = 10
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 1
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
blackmnu.Checked = False
greenmnu.Checked = True
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = False
LightBluemnu.Checked = False
DarkGreenmnu.Checked = False
Yellowmnu.Checked = False
Pinkmnu.Checked = False
End Sub

Private Sub Label4_Click()
Drawcol = 9
Label5.BorderStyle = 0
Label4.BorderStyle = 1
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0

blackmnu.Checked = False
greenmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = True
whitemnu.Checked = False
LightBluemnu.Checked = False
DarkGreenmnu.Checked = False
Yellowmnu.Checked = False
Pinkmnu.Checked = False

End Sub

Private Sub Label5_Click()
Drawcol = 7
Label5.BorderStyle = 1
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
blackmnu.Checked = False
greenmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = True
Yellowmnu.Checked = False
LightBluemnu.Checked = False
DarkGreenmnu.Checked = False
Pinkmnu.Checked = False
End Sub

Private Sub Label6_Click()
Drawcol = 14
Label7.BorderStyle = 0
Label6.BorderStyle = 1
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
blackmnu.Checked = False
greenmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = False
LightBluemnu.Checked = False
Yellowmnu.Checked = True
DarkGreenmnu.Checked = False
Pinkmnu.Checked = False
End Sub

Private Sub Label7_Click()
Drawcol = 11
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label7.BorderStyle = 1
Label6.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
blackmnu.Checked = False
greenmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = False
LightBluemnu.Checked = True
DarkGreenmnu.Checked = False
Yellowmnu.Checked = False
Pinkmnu.Checked = False



End Sub

Private Sub Label8_Click()
Drawcol = 2
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 1
Label9.BorderStyle = 0

blackmnu.Checked = False
greenmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = False
LightBluemnu.Checked = False
Yellowmnu.Checked = False
DarkGreenmnu.Checked = True
Pinkmnu.Checked = False
End Sub

Private Sub Label9_Click()
Drawcol = 13
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 1
blackmnu.Checked = False
greenmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = False
LightBluemnu.Checked = False
Yellowmnu.Checked = False
Pinkmnu.Checked = True
whitemnu.Checked = False
DarkGreenmnu.Checked = False
End Sub

Private Sub LightBluemnu_Click()
Drawcol = 11
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 1
Label8.BorderStyle = 0
Label9.BorderStyle = 0
blackmnu.Checked = False
greenmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = False
LightBluemnu.Checked = True
Yellowmnu.Checked = False
DarkGreenmnu.Checked = False
Pinkmnu.Checked = False

End Sub

Private Sub mnuclear_Click()
Picture1.Cls
End Sub

Private Sub Option1_Click()
Picture1.DrawWidth = 1
End Sub

Private Sub Option2_Click()
Picture1.DrawWidth = 5
End Sub

Private Sub Picture1_Mousemove(button As Integer, shift As Integer, x As Single, y As Single)
If button = 1 Then
Picture1.Line -(x, y), QBColor(Drawcol)
Picture1.CurrentX = x
Picture1.CurrentY = y

End If
'If button = 2 Then Picture1.Line -(x, y), QBColor(9)

Picture1.CurrentX = x
Picture1.CurrentY = y


End Sub
Private Sub Picture1_Mousedown(button As Integer, shift As Integer, x As Single, y As Single)
If button = 2 Then
Picture1.Circle (x, y), 200, QBColor(Drawcol)
End If
End Sub
Sub mnuexit_Click()
End
End Sub

Private Sub Pinkmnu_Click()
Drawcol = 13
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 1
blackmnu.Checked = False
greenmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = False
LightBluemnu.Checked = False
Yellowmnu.Checked = False
DarkGreenmnu.Checked = False
Pinkmnu.Checked = True

End Sub

Private Sub redmnu_Click()
Call GoRed
End Sub

Private Sub whitemnu_Click()
Drawcol = 7
Label5.BorderStyle = 1
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label6.BorderStyle = 0
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
blackmnu.Checked = False
greenmnu.Checked = False
redmnu.Checked = False
bluemnu.Checked = False
whitemnu.Checked = True
LightBluemnu.Checked = False
Yellowmnu.Checked = False
DarkGreenmnu.Checked = False
Pinkmnu.Checked = False
End Sub

Private Sub GoColor(New_Color)
Drawcol = New_Color
Label5.BorderStyle = 0
Label4.BorderStyle = 0
Label3.BorderStyle = 0
Label2.BorderStyle = 0
Label1.BorderStyle = 0
Label6.BorderStyle = 1
Label7.BorderStyle = 0
Label8.BorderStyle = 0
Label9.BorderStyle = 0
blackmnu.Checked = (Drawcol = 0)
greenmnu.Checked = (Drawcol = 10)
redmnu.Checked = (Drawcol = 12)
bluemnu.Checked = (Drawcol = 9)
whitemnu.Checked = (Drawcol = 15)
LightBluemnu.Checked = (Drawcol = 11)
Yellowmnu.Checked = (Drawcol = 14)
DarkGreenmnu.Checked = (Drawcol = 2)
Pinkmnu.Checked = (Drawcol = 13)
End Sub
