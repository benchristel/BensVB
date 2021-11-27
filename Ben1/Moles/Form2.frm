VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Bushwhacker"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Improvements and Graphics by Ben Christel      "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Bushwhacker: Original code written by David Smith"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Unload Form2
End Sub
