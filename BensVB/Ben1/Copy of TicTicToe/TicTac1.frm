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
   Begin VB.CommandButton Command3 
      Caption         =   "Help"
      Height          =   735
      Left            =   360
      TabIndex        =   5
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   735
      Left            =   360
      TabIndex        =   4
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
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   """O"""
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Pick your mark "
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1575
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
'Label1.Caption = CompMar
' Form2.Label1.Caption = "P"
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form4.Visible = True
End Sub

Private Sub Option1_Click()
UserMark = "X"
CompMark = "O"
End Sub

Private Sub Option2_Click()
UserMark = "O"
CompMark = "X"
End Sub
