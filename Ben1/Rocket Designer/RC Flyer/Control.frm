VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form2"
   ScaleHeight     =   2580
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      ToolTipText     =   "Quit"
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdSpecial 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "Special Feature Not Enabled"
      Top             =   480
      Width           =   795
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   720
      TabIndex        =   3
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1320
      TabIndex        =   2
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image imgIdentifier 
      Height          =   1515
      Left            =   2040
      Picture         =   "Control.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Click For Description"
      Top             =   960
      Width           =   1635
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

