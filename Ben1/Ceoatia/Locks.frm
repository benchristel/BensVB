VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLock 
      Caption         =   "closed"
      Height          =   735
      Index           =   9
      Left            =   3480
      TabIndex        =   10
      Top             =   1020
      Width           =   735
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "closed"
      Height          =   735
      Index           =   8
      Left            =   2700
      TabIndex        =   9
      Top             =   1020
      Width           =   735
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "closed"
      Height          =   735
      Index           =   7
      Left            =   1920
      TabIndex        =   8
      Top             =   1020
      Width           =   735
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "closed"
      Height          =   735
      Index           =   6
      Left            =   1140
      TabIndex        =   7
      Top             =   1020
      Width           =   735
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "closed"
      Height          =   735
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   1020
      Width           =   735
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "closed"
      Height          =   735
      Index           =   4
      Left            =   3480
      TabIndex        =   5
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "closed"
      Height          =   735
      Index           =   3
      Left            =   2700
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "closed"
      Height          =   735
      Index           =   2
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "closed"
      Height          =   735
      Index           =   1
      Left            =   1140
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "CLOSED"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   3855
   End
   Begin VB.CommandButton cmdLock 
      Caption         =   "closed"
      Height          =   735
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LockState(0 To 9) As Boolean
Dim LockLink(9, 2) As Integer


Private Sub cmdEnter_Click()
Load Form4
Form4.Visible = True
Unload Form1
End Sub

Private Sub cmdLock_Click(Index As Integer)
Dim i
If LockState(Index) = False Then
For i = 0 To 2
If LockLink(Index, i) > -1 Then
LockState(LockLink(Index, i)) = False
cmdLock(LockLink(Index, i)).Caption = "closed"
End If
Next i
LockState(Index) = True
cmdLock(Index).Caption = "OPEN"
Else
LockState(Index) = False
cmdLock(Index).Caption = "closed"

End If
For i = 0 To 9
If LockState(i) = False Then GoTo 1
Next i
cmdEnter.Caption = "OPEN"
1:
End Sub

Private Sub Form_Load()
LockLink(0, 0) = 4
LockLink(0, 1) = 5
LockLink(0, 2) = 1
LockLink(1, 0) = 5
LockLink(1, 1) = -1
LockLink(1, 2) = -1
LockLink(2, 0) = 7
LockLink(2, 1) = 9
LockLink(2, 2) = 0
LockLink(3, 0) = 9
LockLink(3, 1) = 5
LockLink(3, 2) = 8
LockLink(4, 0) = 6
LockLink(4, 1) = -1
LockLink(4, 2) = 5
LockLink(5, 0) = -1
LockLink(5, 1) = -1
LockLink(5, 2) = -1
LockLink(6, 0) = 1
LockLink(6, 1) = 1
LockLink(6, 2) = 1
LockLink(7, 0) = 9
LockLink(7, 1) = 0
LockLink(7, 2) = 5
LockLink(8, 0) = 0
LockLink(8, 1) = 4
LockLink(8, 2) = 6
LockLink(9, 0) = 8
LockLink(9, 1) = 6
LockLink(9, 2) = 1
End Sub
