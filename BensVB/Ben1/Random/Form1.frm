VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Text            =   "Enter Max"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Text            =   "Enter Min"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Random"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   3240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Max
Dim Min, X
Private Sub Command1_Click()
X = Rnd * (Max - Min + 1)
X = Int(X) + Min
Label1.Caption = X

End Sub

Private Sub Text1_Change()
Min = Text1.Text
End Sub

Private Sub Text2_Change()
Max = Text2.Text
End Sub
