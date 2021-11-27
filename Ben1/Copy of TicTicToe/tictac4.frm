VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3540
   LinkTopic       =   "Form4"
   ScaleHeight     =   4350
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Done"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "tictac4.frx":0000
      Top             =   480
      Width           =   3255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Visible = False
End Sub

