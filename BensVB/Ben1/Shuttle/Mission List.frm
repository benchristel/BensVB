VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mission List"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14790
   FontTransparent =   0   'False
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   14790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Picture..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mission List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14775
   End
   Begin VB.Image imgBack 
      Height          =   6855
      Left            =   0
      Picture         =   "Mission List.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Mission List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Form2.Visible = False
Form1.Visible = True
End Sub

Private Sub Command1_Click()
Form3.Show 1
End Sub

Private Sub Label2_Click()

End Sub
