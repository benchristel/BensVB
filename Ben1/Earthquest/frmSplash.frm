VERSION 5.00
Begin VB.Form frmSplash 
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4320
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   4440
      TabIndex        =   1
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label lblPlay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   840
      TabIndex        =   0
      Top             =   3540
      Width           =   1095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblCancel_Click()
Unload Me
End
End Sub

Private Sub lblPlay_Click()
Load frmLoadout
frmLoadout.Visible = True
Unload Me
End Sub
