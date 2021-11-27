VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Settlers of the New World <> Monday, May 1, 1670"
   ClientHeight    =   5490
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuild 
      BackColor       =   &H000080FF&
      Caption         =   "Build"
      Height          =   735
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1620
      Width           =   735
   End
   Begin VB.Frame fraOptions 
      Height          =   2895
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   3000
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   2460
         Width           =   975
      End
      Begin VB.CommandButton cmdMoreInfo 
         Caption         =   "More Information..."
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2460
         Width           =   1755
      End
      Begin VB.TextBox txtOpt2 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdOpt2 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1755
      End
      Begin VB.CommandButton cmdDone 
         Caption         =   "Done"
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   2460
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtOpt1 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdOpt1 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label lblInfo2 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lblInfo1 
         Height          =   255
         Left            =   1980
         TabIndex        =   10
         Top             =   240
         Width           =   915
      End
   End
   Begin VB.Timer tmrTime 
      Interval        =   5000
      Left            =   3360
      Top             =   5040
   End
   Begin VB.CommandButton cmdFarming 
      BackColor       =   &H0000FF00&
      Caption         =   "Farming"
      Height          =   735
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   735
   End
   Begin VB.CommandButton cmdLogging 
      BackColor       =   &H000040C0&
      Caption         =   "Logging"
      Height          =   735
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblHouses 
      Caption         =   "Houses: 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6780
      TabIndex        =   18
      Top             =   2460
      Width           =   2175
   End
   Begin VB.Label lblTrees 
      Caption         =   "Trees: 5000 "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6780
      TabIndex        =   16
      Top             =   1980
      Width           =   2175
   End
   Begin VB.Label lblWood 
      Caption         =   "Wood: 0 Cords"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6780
      TabIndex        =   15
      Top             =   1500
      Width           =   2175
   End
   Begin VB.Label lblHappy 
      Caption         =   "Happiness: 50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6780
      TabIndex        =   14
      Top             =   540
      Width           =   2175
   End
   Begin VB.Label lblFood 
      Caption         =   "Food: 200 Pounds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6780
      TabIndex        =   13
      Top             =   1020
      Width           =   2175
   End
   Begin VB.Label lblPopulation 
      Caption         =   "Population: 50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6780
      TabIndex        =   12
      Top             =   60
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuSpeed 
      Caption         =   "Speed"
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSlow 
         Caption         =   "Slow"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuMedium 
         Caption         =   "Medium"
         Checked         =   -1  'True
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuFast 
         Caption         =   "Fast"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom"
         Shortcut        =   ^Z
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuild_Click()
Form2.Show
End Sub

Private Sub cmdClose_Click()
fraOptions.Visible = False
End Sub

Private Sub cmdDone_Click()
Dim Change
cmdOpt1.Enabled = True
cmdOpt2.Enabled = True
    Select Case Operation
        Case Is = "Farm"
        Select Case Suboperation
            Case Is = "Farmland"
            AcresFarm = Val(txtOpt1.Text)
            lblInfo1.Caption = "Acres: " & AcresFarm
            txtOpt1.Visible = False
            Case Is = "Farmers"
            txtOpt2.Text = Val(txtOpt2.Text)
            Change = Val(txtOpt2.Text) - Farmers
                If Change > Available Then
                MsgBox "Can't put more people to work than are available."
                Exit Sub
                End If
            Farmers = Val(txtOpt2.Text)
            Available = Available - Change
        End Select
            lblInfo2.Caption = "Farmers: " & Farmers
            With txtOpt2
            .Visible = False
            .Text = ""
            End With
            With txtOpt1
            .Visible = False
            .Text = ""
            End With
            Case Is = "Log"
     Select Case Suboperation
        Case Is = "Area Logged"
        AcresLog = Val(txtOpt1.Text)
        lblInfo1.Caption = "Acres: " & AcresLog
        txtOpt1.Visible = False
        Case Is = "Loggers"
'        OtherWorkers = Farmers
'        txtOpt2.Text = Val(txtOpt2.Text)
        Change = Val(txtOpt2.Text) - Loggers
        txtOpt2.Visible = False
            If Change > Available Then
            MsgBox "Can't put more people to work than are available"
            Exit Sub
            End If
        Loggers = Val(txtOpt2.Text)
        'If Loggers < 5 Then
        'MsgBox ""
        Available = Available - Change
        lblInfo2.Caption = "Loggers: " & Loggers
        With txtOpt2
        .Visible = False
        .Text = ""
        End With
        With txtOpt1
        .Visible = False
        .Text = ""
        End With
    End Select
    End Select
cmdDone.Visible = False
cmdClose.Visible = True
End Sub


Private Sub cmdFarming_Click()
fraOptions.Caption = "Farming"
Operation = "Farm"
lblInfo1.Caption = "Acres: " & AcresFarm
lblInfo2.Caption = "Farmers: " & Farmers
cmdOpt1.Caption = "Change Land Usage"
cmdOpt2.Caption = "Change Workforce"
fraOptions.Visible = True
End Sub

Private Sub cmdLogging_Click()
fraOptions.Caption = "Logging"
Operation = "Log"
lblInfo1.Caption = "Acres: " & AcresLog
lblInfo2.Caption = "Loggers: " & Loggers
cmdOpt1.Caption = "Change Land Usage"
cmdOpt2.Caption = "Change Workforce"
fraOptions.Visible = True
End Sub

Private Sub cmdMoreInfo_Click()
MsgBox ("listindex is " & Form2.lstBuilding.ListIndex)
MsgBox ("item is " & Form2.lstBuilding.Text)
End Sub

Private Sub cmdOpt1_Click()
cmdDone.Visible = True
cmdClose.Visible = False
cmdOpt2.Enabled = False
        txtOpt1.Visible = True
        txtOpt1.SetFocus
    Select Case Operation
        Case Is = "Farm"
        Suboperation = "Farmland"
        Case Is = "Log"
        Suboperation = "Area Logged"
    End Select
End Sub

Private Sub cmdOpt2_Click()
cmdDone.Visible = True
cmdClose.Visible = False
cmdOpt1.Enabled = False
        txtOpt2.Visible = True
        txtOpt2.SetFocus
    Select Case Operation
        Case Is = "Farm"
        Suboperation = "Farmers"
        Case Is = "Log"
        Suboperation = "Loggers"
    End Select
End Sub

Private Sub Form_Load()
Randomize
AcresFarm = 0
Farmers = 0
Loggers = 0
Population = 50
Available = 50
Food = 200
Workers = 0
Happy = 50
Logs = 0
Trees = 5000
Land = 2000
Day = 1
Month = "May"
Year = 1670
Weekday = "Monday"
Form1.Caption = "Settlers of the New World <" & ColonyName & "> " & Weekday & ", " & Month & " " & Day & ", " & Year
Form3.Show
End Sub

Private Sub mnuFast_Click()
With tmrTime
.Enabled = False
.Interval = 2000
.Enabled = True
End With
mnuSlow.Checked = False
mnuMedium.Checked = False
mnuFast.Checked = True
End Sub

Private Sub mnuMedium_Click()
With tmrTime
.Enabled = False
.Interval = 5000
.Enabled = True
End With
mnuSlow.Checked = False
mnuMedium.Checked = True
mnuFast.Checked = False
End Sub

Private Sub mnuOpen_Click()
OpenProject.Show
OpenProject.ZOrder
End Sub

Private Sub mnuPause_Click()
tmrTime.Enabled = False
End Sub

Private Sub mnuQuit_Click()
End
End Sub

Private Sub mnuSlow_Click()
With tmrTime
.Enabled = False
.Interval = 10000
.Enabled = True
End With
mnuSlow.Checked = True
mnuMedium.Checked = False
mnuFast.Checked = False
End Sub

Private Sub mnuZoom_Click()
With tmrTime
.Enabled = False
.Interval = 500
.Enabled = True
End With
End Sub

Private Sub tmrTime_Timer()
Dim TreeGain
Dim ForestLand
Day = Day + 1
'Week = Week + 1
'If Week = 5 Then
'Week = 1
'Month = Month + 1
'End If
'If Month = 13 Then
'Month = 1
'Year = Year + 1
'End If
'If Day = 7 Then
'Form1.Caption = "Settlers of the New World <" & ColonyName & "> Day " & Day & " (Sunday); Week " & Week & "; Month " & Month & ", " & Year
'Else

Select Case Weekday
Case Is = "Monday"
Weekday = "Tuesday"
Case Is = "Tuesday"
Weekday = "Wednesday"
Case Is = "Wednesday"
Weekday = "Thursday"
Case Is = "Thursday"
Weekday = "Friday"
Case Is = "Friday"
Weekday = "Saturday"
Case Is = "Saturday"
Weekday = "Sunday"
Case Is = "Sunday"
Weekday = "Monday"
End Select

Select Case Month
Case Is = "January"
If Day > 31 Then
Month = "February"
Day = 1
End If
Case Is = "February"
    If Year Mod 4 <> 0 Then 'Right(Year, 2) = "00" Or Right(Year, 2) = "04" Or Right(Year, 2) = "08" Or Right(Year, 2) = "16" Or Right(Year, 2) = "20" Or Right(Year, 2) = "24" Or Right(Year, 2) = "28" Or "32" Or Right(Year, 2) = "36" Or Right(Year, 2) = "40" Or Right(Year, 2) = "44" Or Right(Year, 2) = "48" Or Right(Year, 2) = "52" Or Right(Year, 2) = "56" Or Right(Year, 2) = "60" Or Right(Year, 2) = "64" Or Right(Year, 2) = "68" Or Right(Year, 2) = "72" Or Right(Year, 2) = "76" Or Right(Year, 2) = "80" Or Right(Year, 2) = "84" Or Right(Year, 2) = "88" Or Right(Year, 2) = "92" Or Right(Year, 2) = "96" Then
        If Day > 29 Then
        Day = 1
        Month = "March"
        End If
    Else
        If Day = 28 Then
        Month = "March"
        End If
End If
Case Is = "March"
If Day > 31 Then
Month = "April"
Day = 1
End If
Case Is = "April"
If Day > 30 Then
Month = "May"
Day = 1
End If
Case Is = "May"
If Day > 31 Then
Month = "June"
Day = 1
End If
Case Is = "June"
If Day > 30 Then
Month = "July"
Day = 1
End If
Case Is = "July"
If Day > 31 Then
Month = "August"
Day = 1
End If
Case Is = "August"
If Day > 31 Then
Month = "September"
Day = 1
End If
Case Is = "September"
If Day > 30 Then
Month = "October"
Day = 1
End If
Case Is = "October"
If Day > 31 Then
Month = "November"
Day = 1
End If
Case Is = "November"
If Day > 30 Then
Month = "December"
Day = 1
End If
Case Is = "December"
If Day > 31 Then
Month = "January"
Day = 1
Year = Year + 1
MsgBox "It's a new year!", , "Holiday"
End If
End Select
Form1.Caption = "Settlers of the New World <" & ColonyName & "> " & Weekday & ", " & Month & " " & Day & ", " & Year
ForestLand = Land - AcresFarm - AcresLog
    Food = Food - Population * 0.5
    If Month <> "November" And Month <> "December" And Month <> "January" And Month <> "February" And Weekday <> "Sunday" Then
    If Farmers <= AcresFarm * 1.5 Then
        Food = Food + Int(0.1 * Farmers * Farmers)
        End If
        If Farmers >= AcresFarm * 1.5 Then
        Food = Food + Int(0.1 * Farmers * AcresFarm)
        Else
        End If
        End If
        If Food > 4500 And StoreNum = 0 Then
        Food = 4500
        End If
        lblFood.ForeColor = QBColor(0)
    If Food <= 0 Then
    Food = 0
    lblFood.ForeColor = QBColor(12)
    Happy = Happy - 5
        If Happy <= 0 Then
        MsgBox "You are overthrown!", , "Revolution!"
        Happy = 0
        lblHappy.ForeColor = QBColor(12)
        End
        End If
        If Happy <= 15 Then
        lblHappy.ForeColor = QBColor(14)
        End If
    End If
    Select Case Loggers
    Case Is >= AcresLog
    Logs = Logs + Int(0.01 * Loggers * AcresLog)
    Trees = Trees - Int(0.05 * Loggers * AcresLog)
    Case Is < AcresLog
    Logs = Logs + Int(0.01 * Loggers * Loggers)
    Trees = Trees - Int(0.05 * Loggers * Loggers)
    End Select
    If Loggers <= 10 Then
    Select Case Loggers
    Case Is >= AcresLog
    Trees = Trees + Int(0.05 * Loggers * AcresLog)
    Case Is < AcresLog
    Trees = Trees + Int(0.05 * Loggers * Loggers)
    End Select
    End If
    lblFood.Caption = "Food: " & Food & " Pounds"
    If Trees <= 0 And Loggers > 0 Then
    Trees = 0
    MsgBox "There are no more trees.  Logging operations will be shut down", vbInformation, "News from logging camps."
    Available = Available + Loggers
    Loggers = 0
    AcresLog = 0
    If Operation = "Log" Then
    lblInfo1.Caption = "Acres: 0"
    lblInfo2.Caption = "Loggers: 0"
    End If
    End If
    Logs = Logs - Houses
    If Logs < 0 Then
    Logs = 0
    End If
    lblTrees.Caption = "Trees: " & Trees
    TreeGain = Int(ForestLand * 0.001)
    Trees = Trees + TreeGain
        If BuildingNow = True Then
            If Food > 0 Or Logs > LogsNeed Then
            BuildTimeElapsed = BuildTimeElapsed + 1
                If BuildTimeElapsed = BuildTime Then
                    Select Case BuildSpeed
                    Case Is = "Build1"
                    Houses = Houses + 1
                    BuildTimeElapsed = 0
                    Built = Built + 1
                    Logs = Logs - LogsNeed
                    Happy = Happy + HappyAdd
                    Case Is = "BuildAll"
                    Houses = Houses + BuildNum
                    Happy = Happy + HappyAdd
                    SickRisk = SickRisk + SickRiskChange
                    Logs = Logs - LogsNeed
                    MsgBox "Building project complete", vbInformation, "News from building site"
                    End Select
                End If
            Else
            MsgBox "Due to insufficient wood or food," & vbCr & "building projects will be delayed", vbInformation, "News from building site"
            End If
        End If
        If Built = BuildNum Then
        BuildingNow = False
        Built = 0
        End If
lblHappy.Caption = "Happiness: " & Happy
lblWood.Caption = "Wood: " & Logs & " Cords"
lblHouses.Caption = "Houses: " & Houses
End Sub
Private Sub txtOpt1_KeyPress(KeyAscii As Integer)
' MsgBox ("You Pressed " & KeyAscii)
    If KeyAscii = 13 Then
    ' The user pressed return
      Call cmdDone_Click
    '
    cmdOpt2.SetFocus
    End If
End Sub
Private Sub txtOpt2_KeyPress(KeyAscii As Integer)
MsgBox ("You Pressed " & KeyAscii)
    'If KeyAscii = 13 Then
    ' The user pressed return
      'Call cmdDone_Click
    '
    'cmdOpt1.SetFocus
    'End If
End Sub

