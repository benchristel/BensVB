VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   4500
      Left            =   360
      ScaleHeight     =   4440
      ScaleWidth      =   7440
      TabIndex        =   0
      Top             =   720
      Width           =   7500
   End
   Begin VB.Label Label1 
      Caption         =   "Do Loops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGo_Click()
Dim X, Y
Picture1.Cls
Picture1.Line (3000, 2000)-(4000, 2500), QBColor(14), BF
Picture1.CurrentX = 0
Picture1.CurrentY = 0
Do
X = Rnd * 7500
Y = Rnd * 4500
Picture1.Line -(X, Y)
Loop Until X > 3000 And X <= 4000 And Y > 2000 And Y < 2500
End Sub

Private Sub Form_Load()
Randomize
End Sub
