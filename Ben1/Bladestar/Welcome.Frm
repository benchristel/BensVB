VERSION 5.00
Begin VB.Form frmSetup 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   15360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   15360
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picScrollText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   3
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1281
      TabIndex        =   24
      Top             =   1560
      Width           =   19215
   End
   Begin VB.PictureBox picScrollText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1281
      TabIndex        =   23
      Top             =   7200
      Width           =   19215
   End
   Begin VB.Timer tmrScroll 
      Interval        =   10
      Left            =   16440
      Top             =   240
   End
   Begin VB.PictureBox picScrollText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1281
      TabIndex        =   22
      Top             =   5400
      Width           =   19215
   End
   Begin VB.PictureBox picScrollText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1281
      TabIndex        =   21
      Top             =   13800
      Width           =   19215
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[ close ]"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   615
      Left            =   8760
      TabIndex        =   18
      Top             =   9240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblInversion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "inversion off"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   615
      Left            =   7320
      TabIndex        =   20
      Top             =   8040
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label lblLayout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "layout: rectangular"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   615
      Left            =   6840
      TabIndex        =   19
      Top             =   7440
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label lblOptions 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[ options ]"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   615
      Left            =   8400
      TabIndex        =   17
      Top             =   9240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblLogout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "||| LOGOUT |||"
      BeginProperty Font 
         Name            =   "Weltron Urban"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   615
      Left            =   8280
      TabIndex        =   16
      Top             =   14040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblDifficulty 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Left            =   6960
      TabIndex        =   15
      ToolTipText     =   "Change Difficulty Level"
      Top             =   8640
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Left            =   6960
      TabIndex        =   14
      ToolTipText     =   "Start Level"
      Top             =   8040
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Left            =   6960
      TabIndex        =   13
      ToolTipText     =   "Delete Account"
      Top             =   7440
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "| [ ] |"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Index           =   9
      Left            =   6960
      TabIndex        =   12
      Top             =   12840
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "| [ ] |"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Index           =   8
      Left            =   6960
      TabIndex        =   11
      Top             =   12240
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblStart 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "sign in"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   8640
      TabIndex        =   10
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "| [ ] |"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Index           =   7
      Left            =   6960
      TabIndex        =   9
      Top             =   11640
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "| [ ] |"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Index           =   6
      Left            =   6960
      TabIndex        =   8
      Top             =   11040
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "| [ ] |"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Index           =   5
      Left            =   6960
      TabIndex        =   7
      Top             =   10440
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "| [ ] |"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Index           =   4
      Left            =   6960
      TabIndex        =   6
      Top             =   9840
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "| [ ] |"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Index           =   3
      Left            =   6960
      TabIndex        =   5
      Top             =   9240
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "| [ ] |"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Index           =   2
      Left            =   6960
      TabIndex        =   4
      Top             =   8640
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "| [ ] |"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Index           =   1
      Left            =   6960
      TabIndex        =   3
      Top             =   8040
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblBladestar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B L A D E S T A R  2"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   18855
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "| [ ] |"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   495
      Index           =   0
      Left            =   6960
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "||| EXIT |||"
      BeginProperty Font 
         Name            =   "Weltron Urban"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00468C0A&
      Height          =   615
      Left            =   8280
      TabIndex        =   0
      Top             =   14640
      Width           =   2535
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MaxLevel = 9
Dim scrollx(0 To 3) As Double, scrollspeed(0 To 3) As Single
Dim ScrollDC(0 To 3) As Long, ScrollBackDC As Long, endit As Boolean
Private Sub Form_Load()
Dim temp, i, j
scrollspeed(0) = 1.5
scrollspeed(1) = 1
scrollspeed(2) = 1.2
scrollspeed(3) = 1.5
Me.BackColor = RGB(10, 30, 10)
lblBladestar.ForeColor = RGB(20, 200, 100)
ScrollDC(0) = GenerateDC(App.Path & "\Graphics\ScrollText2.bmp")
ScrollDC(1) = GenerateDC(App.Path & "\Graphics\ScrollText1.bmp")
ScrollDC(2) = GenerateDC(App.Path & "\Graphics\ScrollText1.bmp")
ScrollDC(3) = GenerateDC(App.Path & "\Graphics\ScrollText2.bmp")

ScrollBackDC = GenerateDC(App.Path & "\Graphics\ScrollBackground.bmp")
'read account data from file
'9 lines
'name
'level
'difficulty (1 = easy, 2 = med, 3 = hard, 4 = epic)
'weapons 1
'ammo 1
'clipammo 1
'same for 2
Open App.Path & "\Data\Accounts.dat" For Input As #1
For i = 0 To 9
Line Input #1, temp
AccountData(i).Name = temp
Line Input #1, temp
AccountData(i).Level = temp
Line Input #1, temp
AccountData(i).Difficulty = temp
For j = 1 To 2
Line Input #1, temp
AccountData(i).Weapon(j) = temp
Line Input #1, temp
AccountData(i).Ammo(j) = temp
Line Input #1, temp
AccountData(i).ClipAmmo(j) = temp
Next j
Line Input #1, temp
AccountData(i).Layout = temp
Line Input #1, temp
AccountData(i).Inverted = temp
If AccountData(i).Name <> "" Then lblLogin(i).Caption = AccountData(i).Name
Next i
Close #1
endit = False
Call TimerLoop
End Sub


Private Sub lblClose_Click()
lblClose.Visible = False
lblName.Visible = True
lblLevel.Visible = True
lblDifficulty.Visible = True
lblLayout.Visible = False
lblInversion.Visible = False
lblOptions.Visible = True
End Sub

Private Sub lblDifficulty_Click()
Select Case AccountData(CurrentAccount).Difficulty
Case Is = 1
    AccountData(CurrentAccount).Difficulty = 2
    lblDifficulty.Caption = "medium"
Case Is = 2
    AccountData(CurrentAccount).Difficulty = 3
    lblDifficulty.Caption = "hard"
Case Is = 3
    AccountData(CurrentAccount).Difficulty = 4
    lblDifficulty.Caption = "epic"
Case Is = 4
    AccountData(CurrentAccount).Difficulty = 1
    lblDifficulty.Caption = "easy"
End Select
End Sub

Private Sub lblExit_Click()
Dim i, j
'save account data
Open App.Path & "\Data\Accounts.dat" For Output As #1
For i = 0 To 9
Print #1, AccountData(i).Name
Print #1, AccountData(i).Level
Print #1, AccountData(i).Difficulty
For j = 1 To 2
Print #1, AccountData(i).Weapon(j)
Print #1, AccountData(i).Ammo(j)
Print #1, AccountData(i).ClipAmmo(j)
Next j
Print #1, AccountData(i).Layout
Print #1, AccountData(i).Inverted
Next i
Close #1
'=====
Unload frmSetup
endit = True
End
End Sub

Private Sub lblInversion_Click()
Select Case AccountData(CurrentAccount).Inverted
Case Is = True
    lblInversion.Caption = "inversion: off"
    AccountData(CurrentAccount).Inverted = False
Case Is = False
    lblInversion.Caption = "inversion: on"
    AccountData(CurrentAccount).Inverted = True
End Select
End Sub

Private Sub lblLayout_Click()
Select Case AccountData(CurrentAccount).Layout
Case Is = 1
    AccountData(CurrentAccount).Layout = 2
    lblLayout.Caption = "layout: bladestar classic"
Case Is = 2
    AccountData(CurrentAccount).Layout = 3
    lblLayout.Caption = "layout: linear"
Case Is = 3
    AccountData(CurrentAccount).Layout = 1
    lblLayout.Caption = "layout: rectangular"
End Select
End Sub

Private Sub lblLevel_Click()
Dim i
Call InitializeWeapons
Select Case AccountData(CurrentAccount).Difficulty
Case Is = 1 'easy
PlayerBonus = 1.2
EnemyBonus = 0.8
Case Is = 2 'med
PlayerBonus = 1
EnemyBonus = 1
Case Is = 3 'hard
PlayerBonus = 0.8
EnemyBonus = 1.1
Case Is = 4 'epic
PlayerBonus = 0.7
EnemyBonus = 1.2
End Select
'==========
'load level
'==========
For i = 1 To 2
Player.Weapon(i) = AccountData(CurrentAccount).Weapon(i)
Player.Ammo(i) = AccountData(CurrentAccount).Ammo(i)
Player.ClipAmmo(i) = AccountData(CurrentAccount).ClipAmmo(i)
Next i
Select Case AccountData(CurrentAccount).Level
Case Is = 1
Call LoadLevel1
Case Is = 2
Call LoadLevel2
Case Is = 3
Call LoadLevel3
Case Is = 4
Call LoadLevel4
Case Is = 5
Call LoadLevel5
Case Is = 6
Call LoadLevel6
Case Is = 7
Call LoadLevel7
Case Is = 8
Call LoadLevel8
Case Is = 9
Call LoadLevel9
End Select


lblName.Enabled = False
lblLevel.Enabled = False
lblDifficulty.Enabled = False
On Error Resume Next
frmMain.Visible = True
lblName.Enabled = True
If AccountData(CurrentAccount).Level > MaxLevel Then
lblLevel.Caption = "complete"
Else
lblLevel.Caption = "level " & AccountData(CurrentAccount).Level
lblLevel.Enabled = True
End If
lblDifficulty.Enabled = True

End Sub

Private Sub lblLogin_Click(Index As Integer)
Dim i
CurrentAccount = Index
If AccountData(Index).Name = "" Then
    AccountData(Index).Name = InputBox("Enter a name for this account.", "New Account")
    If Len(AccountData(Index).Name) > 15 Then
    MsgBox "Please enter a name 15 characters or shorter", , "Error"
    AccountData(Index).Name = ""
    Exit Sub
    End If
    If AccountData(Index).Name = "" Then
    MsgBox "Please enter a name between 1 and 15 characters long.", , "Error"
    Exit Sub
    End If
End If
    CurrentLevel = AccountData(Index).Level
With Player
    .Ammo(1) = AccountData(Index).Ammo(1)
    .ClipAmmo(1) = AccountData(Index).ClipAmmo(1)
    .Weapon(1) = AccountData(Index).Weapon(1)
    .Ammo(2) = AccountData(Index).Ammo(2)
    .ClipAmmo(2) = AccountData(Index).ClipAmmo(2)
    .Weapon(2) = AccountData(Index).Weapon(2)
End With
For i = 0 To 9
    lblLogin(i).Visible = False
Next i
Select Case AccountData(Index).Difficulty
Case Is = 1
    lblDifficulty.Caption = "easy"
Case Is = 2
    lblDifficulty.Caption = "medium"
Case Is = 3
    lblDifficulty.Caption = "hard"
Case Is = 4
    lblDifficulty.Caption = "epic"
End Select
Select Case AccountData(CurrentAccount).Layout
Case Is = 1
    lblLayout.Caption = "layout: rectangular"
Case Is = 2
    lblLayout.Caption = "layout: bladestar classic"
Case Is = 3
    lblLayout.Caption = "layout: linear"
End Select
Select Case AccountData(CurrentAccount).Inverted
Case Is = True
    lblInversion.Caption = "inversion: on"
Case Is = False
    lblInversion.Caption = "inversion: off"
End Select
    lblName.Caption = AccountData(Index).Name
    lblName.Visible = True
If AccountData(Index).Level <= MaxLevel Then
    lblLevel.Caption = "level " & AccountData(Index).Level
    lblLevel.Enabled = True
Else
    lblLevel.Caption = "complete"
    lblLevel.Enabled = False
End If
    lblLevel.Visible = True
    lblDifficulty.Visible = True
    lblLogout.Visible = True
    lblOptions.Visible = True
End Sub

Private Sub lblLogout_Click()
Dim i
lblName.Visible = False
lblLevel.Visible = False
lblDifficulty.Visible = False
lblLogout.Visible = False
lblOptions.Visible = False
For i = 0 To 9
    If AccountData(i).Name <> "" Then
    lblLogin(i).Caption = AccountData(i).Name
    Else
    lblLogin(i).Caption = "| [ ] |"
    End If
    lblLogin(i).Visible = True
Next i
End Sub

Private Sub lblName_Click()
Dim i
If MsgBox("Do you want to delete this account?", vbYesNo, "Delete Account") = vbYes Then
    With AccountData(CurrentAccount)
    For i = 1 To 2
        .Ammo(i) = 0
        .ClipAmmo(i) = 0
        .Weapon(i) = 0
    Next i
        .Difficulty = 1
        .Level = 1
        .Name = ""
    End With
    lblName.Visible = False
    lblLevel.Visible = False
    lblDifficulty.Visible = False
    lblLogout.Visible = False
    lblOptions.Visible = False
    For i = 0 To 9
        If AccountData(i).Name <> "" Then
            lblLogin(i).Caption = AccountData(i).Name
        Else
            lblLogin(i).Caption = "| [ ] |"
        End If
        lblLogin(i).Visible = True
    Next i
End If
End Sub

Private Sub lblOptions_Click()
lblClose.Visible = True
lblName.Visible = False
lblLevel.Visible = False
lblDifficulty.Visible = False
lblLayout.Visible = True
lblInversion.Visible = True
lblOptions.Visible = False
End Sub

Private Sub lblStart_Click()
Dim i
lblStart.Visible = False
For i = 0 To 9
    lblLogin(i).Visible = True
Next i
End Sub

Private Sub TimerLoop()
Dim FrameCount As Integer, i
Dim CurrentTick As Long
Dim LastTick As Long
Const FrameDifference As Long = 10
Const FPS = 100
FrameCount = 0
Me.Show
Do
If endit = True Then Exit Do
    CurrentTick = GetTickCount()
    If CurrentTick - LastTick > FrameDifference Then
For i = 0 To 3
scrollx(i) = scrollx(i) + scrollspeed(i)
If scrollx(i) > 1500 Then scrollx(i) = scrollx(i) - 1500
'BitBlt picScrollText(i).hdc, 0, 0, 1500, 13, ScrollBackDC, 0, 0, vbSrcCopy
BitBlt picScrollText(i).hdc, scrollx(i), 0, 1500, 18, ScrollDC(i), 0, 0, vbSrcCopy
BitBlt picScrollText(i).hdc, scrollx(i) - 1500, 0, 1500, 18, ScrollDC(i), 0, 0, vbSrcCopy
picScrollText(i).Refresh
Next i
DoEvents
LastTick = CurrentTick
    Else
        DoEvents
    End If

Loop
End Sub
