VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton button2 
      Caption         =   "click me"
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton EndButton 
      Caption         =   "End"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton button1 
      Caption         =   "push me"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   1875
      Left            =   4080
      Picture         =   "ben1.frx":0000
      Stretch         =   -1  'True
      Top             =   4200
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   4860
      Left            =   600
      Picture         =   "ben1.frx":1AC2
      Top             =   2520
      Visible         =   0   'False
      Width           =   8370
   End
   Begin VB.Label Output1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Not Working"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub button1_Click()
Output1.Caption = "It Worked."
Image1.Visible = True
Image2.Visible = False
End Sub

Private Sub button2_Click()
Output1.Caption = "Wow."
Image1.Visible = True
Image2.Visible = True
End Sub

Private Sub EndButton_Click()
End
End Sub

Private Sub Image1_Click()
Output1.Caption = "The Plane."
Image2.Visible = False
End Sub

