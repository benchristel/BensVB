VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Picture Select"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Just an orbiter."
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   1800
      Width           =   1470
   End
   Begin VB.Label lbl2Description 
      Caption         =   "A space shuttle in the air."
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   1470
   End
   Begin VB.Label lbl1Description 
      Caption         =   "A space shuttle on the pad."
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   1470
   End
   Begin VB.Image imgShuttle 
      Height          =   735
      Left            =   120
      Picture         =   "Picture Select.frx":0000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image imgShuttleInFlight 
      Height          =   735
      Left            =   120
      Picture         =   "Picture Select.frx":6EFC
      Stretch         =   -1  'True
      Top             =   960
      Width           =   735
   End
   Begin VB.Image imgOnThePad 
      Height          =   735
      Left            =   120
      Picture         =   "Picture Select.frx":22D5D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
Unload Form3
End Sub

Private Sub Form_Load()

End Sub
