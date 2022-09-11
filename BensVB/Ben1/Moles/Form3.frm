VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Moles Instructions"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   5460
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "But if, instead of going to ""Help"" you went down to ""Exit Game"", the game would  end , and the main screen would dissapear. "
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   4575
   End
   Begin VB.Label Label7 
      Caption         =   $"Form3.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "Menus:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   $"Form3.frx":00EA
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   $"Form3.frx":01A6
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "Hits And Misses:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Startup:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "When you see the main screen, click  on the ""Go"" button to start Bush."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Unload Form3
End Sub

