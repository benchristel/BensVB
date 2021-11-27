VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6540
   ClientLeft      =   4425
   ClientTop       =   2160
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   6540
   Begin VB.Frame Frame2 
      Caption         =   "Pick a font"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
      Begin VB.OptionButton Option4 
         Caption         =   "Comic Sans MS"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "MS Sans Serif"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   """X"""
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1320
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   """O"""
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pick your mark"
      Height          =   1335
      Left            =   2280
      TabIndex        =   5
      Top             =   1080
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Load Form3
Form3.Show
Form1.Visible = False
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form4.Visible = True
End Sub

Private Sub Form_Load()
UserMark = "X"
CompMark = "O"
End Sub

Private Sub Option1_Click()
UserMark = "X"
CompMark = "O"
End Sub

Private Sub Option2_Click()
UserMark = "O"
CompMark = "X"
End Sub

Private Sub Option3_Click()
    With Form3.Label1(1).Font
        .Name = ""
        .Bold = False
        .Size = 48

End Sub

Private Sub Option4_Click()
    With Form3.Label1(1).Font
        .Name = "Comic Sans MS"
        .Bold = False
        .Size = 48
'        .Strikethrough = False
'        .Superscript = False
'        .Subscript = False
'        .OutlineFont = False
'        .Shadow = False
'        .Underline = xlUnderlineStyleNone
'        .ColorIndex = xlAutomatic
    End With
End Sub
