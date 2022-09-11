VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   LinkTopic       =   "Form3"
   ScaleHeight     =   6210
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   495
      Left            =   5040
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "&Restart"
      Height          =   495
      Left            =   5040
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Click on a Square"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      TabIndex        =   9
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   8
      Left            =   3720
      TabIndex        =   8
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   7
      Left            =   2400
      TabIndex        =   7
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   6
      Left            =   1080
      TabIndex        =   6
      Top             =   3480
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   5
      Left            =   3720
      TabIndex        =   5
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   4
      Left            =   2400
      TabIndex        =   4
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   1
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   1200
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOptions_Click()
Form1.Visible = True
Form3.Visible = False
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub cmdRestart_Click()
    For i = 0 To 8
     Label1.Item(i).Caption = ""
    Label10.Caption = "Click on a Square"
    Next i

End Sub

Private Sub Label1_Click(Index As Integer)
If Label1.Item(Index).Caption <> Empty Then Exit Sub
Label1.Item(Index).Caption = UserMark
Call CheckUserWin
Call CheckCompwin
Call checkblock
Call CheckSquare
End Sub
Private Sub checkblock()
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 0 To 2
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if need to block user
'
 If cct = 0 And uct = 2 Then
 For i = 0 To 2
    If Label1.Item(i).Caption = "" Then
        Label1.Item(i).Caption = CompMark
        Exit Sub
    End If
  Next i
End If
' row2
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 3 To 5
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if need to block user
'
 If cct = 0 And uct = 2 Then
 For i = 3 To 5
    If Label1.Item(i).Caption = "" Then
        Label1.Item(i).Caption = CompMark
        Exit Sub
    End If
  Next i
End If ' row3
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 6 To 8
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if need to block user
'
 If cct = 0 And uct = 2 Then
 For i = 6 To 8
    If Label1.Item(i).Caption = "" Then
        Label1.Item(i).Caption = CompMark
        Exit Sub
    End If
  Next i
End If
'column1
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 0 To 6 Step 3
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if need to block user
'
 If cct = 0 And uct = 2 Then
 For i = 0 To 6 Step 3
    If Label1.Item(i).Caption = "" Then
        Label1.Item(i).Caption = CompMark
        Exit Sub
    End If
  Next i
End If
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 1 To 7 Step 3
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if need to block user
'
 If cct = 0 And uct = 2 Then
 For i = 1 To 7 Step 3
    If Label1.Item(i).Caption = "" Then
        Label1.Item(i).Caption = CompMark
        Exit Sub
    End If
  Next i
End If
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 2 To 8 Step 3
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if need to block user
'
 If cct = 0 And uct = 2 Then
 For i = 2 To 8 Step 3
    If Label1.Item(i).Caption = "" Then
        Label1.Item(i).Caption = CompMark
        Exit Sub
    End If
  Next i
End If
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 0 To 8 Step 4
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if need to block user
'
 If cct = 0 And uct = 2 Then
 For i = 0 To 8 Step 4
    If Label1.Item(i).Caption = "" Then
        Label1.Item(i).Caption = CompMark
        Exit Sub
    End If
  Next i
End If
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 2 To 6 Step 2
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if need to block user
'
 If cct = 0 And uct = 2 Then
 For i = 2 To 6 Step 2
    If Label1.Item(i).Caption = "" Then
        Label1.Item(i).Caption = CompMark
        Exit Sub
    End If
  Next i
End If
        If Label1.Item(4).Caption <> Empty Then GoTo ck2
        Label1.Item(4).Caption = CompMark
        Exit Sub
ck2:
        If Label1.Item(2).Caption <> Empty Then GoTo ck0
        Label1.Item(2).Caption = CompMark
        Exit Sub
ck0:
        If Label1.Item(0).Caption <> Empty Then GoTo ck6
        Label1.Item(0).Caption = CompMark
        Exit Sub
ck6:
        If Label1.Item(6).Caption <> Empty Then GoTo ck8
        Label1.Item(6).Caption = CompMark
        Exit Sub
ck8:
        If Label1.Item(8).Caption <> Empty Then GoTo ck1
        Label1.Item(8).Caption = CompMark
        Exit Sub
ck1:
        If Label1.Item(1).Caption <> Empty Then GoTo ck3
        Label1.Item(1).Caption = CompMark
        Exit Sub
ck3:
        If Label1.Item(3).Caption <> Empty Then GoTo ck5
        Label1.Item(3).Caption = CompMark
        Exit Sub
ck5:
        If Label1.Item(5).Caption <> Empty Then GoTo ck7
        Label1.Item(5).Caption = CompMark
        Exit Sub
ck7:
        If Label1.Item(7).Caption <> Empty Then Exit Sub
        Label1.Item(7).Caption = CompMark
        Label10.Caption = "It's a draw!"
End Sub
Private Sub CheckUserWin()
'uct = 0  ' number of user marks found
    For i = 0 To 2
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
'  Check for user win

   If uct = 3 Then GoTo UserWins 'if
   uct = 0  ' number of user marks found
    For i = 3 To 5
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
'  Check for user win
'
   If uct = 3 Then GoTo UserWins 'if
   '
   ' Third Row
   '
   uct = 0  ' number of user marks found
    For i = 6 To 8
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
'  Check for user win
'
   If uct = 3 Then GoTo UserWins 'if
'
' Column One
'
   uct = 0  ' number of user marks found
    For i = 0 To 6 Step 3
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
'  Check for user win
'
   If uct = 3 Then GoTo UserWins 'if
   
   uct = 0  ' number of user marks found
    For i = 1 To 7 Step 3
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
'  Check for user win
'
   If uct = 3 Then GoTo UserWins 'if
   
   uct = 0  ' number of user marks found
    For i = 2 To 8 Step 3
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
'  Check for user win
'
   If uct = 3 Then GoTo UserWins 'if
   uct = 0  ' number of user marks found
    For i = 0 To 8 Step 4
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
'  Check for user win
'
   If uct = 3 Then GoTo UserWins 'if
   
   
   uct = 0  ' number of user marks found
    For i = 2 To 6 Step 2
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
'  Check for user win
'
   If uct = 3 Then GoTo UserWins 'if
 Exit Sub
   
UserWins:
   Label10.Caption = "You won!"
      Call EndOfGame
   Exit Sub
   
   
End Sub
Private Sub CheckCompwin()
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 0 To 2
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
'  Check for computer win
'
   If cct = 2 And uct = 0 Then
      Label1.Item(0).Caption = CompMark
      Label1.Item(1).Caption = CompMark
      Label1.Item(2).Caption = CompMark
      Label10.Caption = "I won"
      Call EndOfGame
    End If
' row2
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 3 To 5
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if computer can win
'
   If cct = 2 And uct = 0 Then
      Label1.Item(3).Caption = CompMark
      Label1.Item(4).Caption = CompMark
      Label1.Item(5).Caption = CompMark
      Label10.Caption = "I won"
            Call EndOfGame

    End If
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 6 To 8
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if computer can win
'
   If cct = 2 And uct = 0 Then
      Label1.Item(6).Caption = CompMark
      Label1.Item(7).Caption = CompMark
      Label1.Item(8).Caption = CompMark
      Label10.Caption = "I won"
            Call EndOfGame

    End If
'column1
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 0 To 6 Step 3
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if computer can win
'
   If cct = 2 And uct = 0 Then
      Label1.Item(0).Caption = CompMark
      Label1.Item(3).Caption = CompMark
      Label1.Item(6).Caption = CompMark
      Label10.Caption = "I won"
            Call EndOfGame

    End If
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 1 To 7 Step 3
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if computer can win
'
   If cct = 2 And uct = 0 Then
      Label1.Item(1).Caption = CompMark
      Label1.Item(4).Caption = CompMark
      Label1.Item(7).Caption = CompMark
      Label10.Caption = "I won"
      Call EndOfGame
    End If
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 2 To 8 Step 3
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if computer can win
'
   If cct = 2 And uct = 0 Then
      Label1.Item(2).Caption = CompMark
      Label1.Item(5).Caption = CompMark
      Label1.Item(8).Caption = CompMark
      Label10.Caption = "I won"
      Call EndOfGame
    End If
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 0 To 8 Step 4
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if computer can win
'
   If cct = 2 And uct = 0 Then
      Label1.Item(0).Caption = CompMark
      Label1.Item(4).Caption = CompMark
      Label1.Item(8).Caption = CompMark
      Label10.Caption = "I won"
      Call EndOfGame
    End If
uct = 0  ' number of user marks found
cct = 0  ' number of computer marks found
    For i = 2 To 6 Step 2
    If Label1.Item(i).Caption = CompMark Then cct = cct + 1
    If Label1.Item(i).Caption = UserMark Then uct = uct + 1
    Next i
'
' Check to see if computer can win
'
   If cct = 2 And uct = 0 Then
      Label1.Item(2).Caption = CompMark
      Label1.Item(4).Caption = CompMark
      Label1.Item(6).Caption = CompMark
      Label10.Caption = "I won"
      Call EndOfGame
    End If

End Sub
Sub CheckSquare()



        End Sub
        

Sub EndOfGame()
    For i = 0 To 8
     Label1.Item(i).Enabled = False
    
    Next i
    cmdOptions.Visible = True
    cmdQuit.Visible = True

End Sub
