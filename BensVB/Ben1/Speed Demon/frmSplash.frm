VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmSplash.frx":030A
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPause 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image5 
      Height          =   900
      Left            =   4980
      Picture         =   "frmSplash.frx":0614
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   900
   End
   Begin VB.Image Image4 
      Height          =   900
      Left            =   3900
      Picture         =   "frmSplash.frx":091E
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   900
   End
   Begin VB.Image Image3 
      Height          =   900
      Left            =   2820
      Picture         =   "frmSplash.frx":0C28
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   900
   End
   Begin VB.Image Image2 
      Height          =   900
      Left            =   1740
      Picture         =   "frmSplash.frx":0F32
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   720
      Picture         =   "frmSplash.frx":123C
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   900
   End
   Begin VB.Image imgLogo 
      Height          =   4320
      Left            =   0
      Picture         =   "frmSplash.frx":1546
      Top             =   -60
      Width           =   8640
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub tmrPause_Timer()
Load frmSelectTrack
Unload frmSplash
End Sub
