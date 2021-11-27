VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5325
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrShelf 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1680
      Top             =   1200
   End
   Begin VB.Timer tmrPresents 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1080
      Top             =   1200
   End
   Begin VB.Timer tmrBenese 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   480
      Top             =   1200
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
tmrBenese.Enabled = True
End Sub

Private Sub tmrBenese_Timer()
Label1.Caption = "BENESE"
Label2.Caption = "Industries"
tmrPresents.Enabled = True
tmrBenese.Enabled = False
End Sub

Private Sub tmrPresents_Timer()
Label2.Caption = "Presents"
Label1.Caption = ""
tmrShelf.Enabled = True
tmrPresents.Enabled = False
End Sub

Private Sub tmrShelf_Timer()
Label1.Caption = "Shelf Watcher"
Label2.Caption = ""
tmrShelf.Enabled = False
End Sub
