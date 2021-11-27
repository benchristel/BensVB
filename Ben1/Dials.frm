VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   2280
      X2              =   2280
      Y1              =   1560
      Y2              =   720
   End
   Begin VB.Label lbl00 
      Alignment       =   1  'Right Justify
      Caption         =   "00"
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin VB.Shape shpFrame 
      BackColor       =   &H8000000B&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   2415
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NeedleY

Private Sub Form_Load()
NeedleY = 1560
NeedleY = NeedleY + 20
Line1.Y1 = NeedleY
End Sub
