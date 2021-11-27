VERSION 5.00
Begin VB.Form OpenProject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Project"
   ClientHeight    =   1125
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEnter 
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   4575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblInstructions 
      Caption         =   "Enter the name of the project you wish to open."
      Height          =   435
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "OpenProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
Unload (OpenProject)
End Sub

Private Sub cmdOK_Click()

End Sub

Private Sub txtEnter_Change()

End Sub
