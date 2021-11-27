VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Square Numbers"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go!"
      Height          =   435
      Left            =   2340
      TabIndex        =   2
      Top             =   1080
      Width           =   915
   End
   Begin VB.ListBox List1 
      Height          =   6105
      Left            =   600
      TabIndex        =   0
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Loc 
      Alignment       =   2  'Center
      Caption         =   "Lockers Remaining Open"
      Height          =   495
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Progress:"
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblProgress 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' The matrix LockerState indicates if a locker
' is open or closed.  True = Open
Dim LockerState(1000) As Boolean
Dim LockerNumber
Dim Studentnumber
Private Sub cmdGo_Click()
'
' Each student runs past each locker in turn.
'
For Studentnumber = 1 To 1000
    For LockerNumber = 1 To 1000
    '
    ' As each student runs past each locker, he will change the
    ' state of the locker if the locker number is a multiple of
    ' the student number.
    ' The MOD function is 0 if this is the case.
    '
        If LockerNumber Mod Studentnumber = 0 Then
            LockerState(LockerNumber) = Not (LockerState(LockerNumber))
        End If
    Next LockerNumber
'
' This label updates the output screen as each student
' completes his run through the lockers.
' DoEvents refreshes the screen.
'
lblProgress.Caption = "Student #" & Studentnumber
DoEvents
Next Studentnumber
'
' The state of all lockers is now checked, and
' the lockers that remain open are added to
' the output list.
'
For LockerNumber = 1 To 1000
If LockerState(LockerNumber) = True Then
List1.AddItem (LockerNumber)
End If
Next LockerNumber
'
' End of program
'
End Sub
