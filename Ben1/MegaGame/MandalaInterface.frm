VERSION 5.00
Begin VB.Form frmInterface 
   Caption         =   "Form2"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11715
   LinkTopic       =   "Form2"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   10995
      Left            =   60
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   800.895
      TabIndex        =   0
      Top             =   60
      Width           =   12975
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[empty]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   10
      Left            =   13080
      TabIndex        =   12
      Top             =   2460
      Width           =   2115
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[empty]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   9
      Left            =   13080
      TabIndex        =   11
      Top             =   2700
      Width           =   2115
   End
   Begin VB.Label lblMoney 
      BackColor       =   &H00000000&
      Caption         =   "$0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   13080
      TabIndex        =   10
      Top             =   3060
      Width           =   2115
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[empty]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   8
      Left            =   13080
      TabIndex        =   9
      Top             =   2220
      Width           =   2115
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[empty]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   7
      Left            =   13080
      TabIndex        =   8
      Top             =   1980
      Width           =   2115
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[empty]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   6
      Left            =   13080
      TabIndex        =   7
      Top             =   1620
      Width           =   2115
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[empty]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   13080
      TabIndex        =   6
      Top             =   1380
      Width           =   2115
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[empty]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   13080
      TabIndex        =   5
      Top             =   1140
      Width           =   2115
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[empty]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   13080
      TabIndex        =   4
      Top             =   780
      Width           =   2115
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[empty]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   13080
      TabIndex        =   3
      Top             =   540
      Width           =   2115
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[empty]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   13080
      TabIndex        =   2
      Top             =   300
      Width           =   2115
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[empty]"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   13080
      TabIndex        =   1
      Top             =   60
      Width           =   2115
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgInventory_Click(Index As Integer)

End Sub

Private Sub lblInventory_Click(Index As Integer)

End Sub
