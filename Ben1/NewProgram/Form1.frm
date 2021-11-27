VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   11190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO"
      Height          =   255
      Left            =   9000
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtNotes 
      Height          =   5895
      Left            =   6840
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   1200
      Width           =   4215
   End
   Begin VB.TextBox txtUser 
      Height          =   615
      Left            =   6840
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0015
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblDisplay 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
