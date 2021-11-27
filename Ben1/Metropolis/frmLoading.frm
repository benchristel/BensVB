VERSION 5.00
Begin VB.Form frmLoading 
   BackColor       =   &H00000000&
   ClientHeight    =   2880
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   3165
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   2880
   ScaleWidth      =   3165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgMoonPic 
      Height          =   2880
      Left            =   0
      Picture         =   "frmLoading.frx":0000
      Top             =   0
      Width           =   3150
   End
End
Attribute VB_Name = "frmLoading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Visible = True
Me.ZOrder
Refresh
Load frmMain
frmMain.Visible = True
Unload Me
End Sub
