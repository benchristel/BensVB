VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Building choices"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6885
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   4140
      Width           =   2055
   End
   Begin VB.CommandButton cmdBuildAll 
      Caption         =   "Build &All"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdBuild1 
      Caption         =   "Build &One At a Time"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3300
      Width           =   2055
   End
   Begin VB.TextBox txtBuildNum 
      Enabled         =   0   'False
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   1680
      Width           =   435
   End
   Begin VB.ListBox lstBuilding 
      Height          =   1035
      ItemData        =   "frmBuild.frx":0000
      Left            =   60
      List            =   "frmBuild.frx":0022
      TabIndex        =   1
      Top             =   600
      Width           =   1995
   End
   Begin VB.Label lblInstruct1 
      Alignment       =   2  'Center
      Caption         =   "Build            Buildings"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   1740
      Width           =   1995
   End
   Begin VB.Label lblDescription 
      Height          =   1140
      Left            =   60
      TabIndex        =   2
      Top             =   2100
      Width           =   1935
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBuilding 
      BorderStyle     =   1  'Fixed Single
      Height          =   3930
      Left            =   2100
      Picture         =   "frmBuild.frx":00A0
      Stretch         =   -1  'True
      Top             =   600
      Width           =   4740
   End
   Begin VB.Label lblBuilding 
      Alignment       =   2  'Center
      Caption         =   "Choose a Building"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdBuild1_Click()
BuildSpeed = "Build1"
BuildingNow = True
Call Form_Load
BuildTimeElapsed = 0
End Sub

Private Sub cmdBuildAll_Click()
BuildSpeed = "BuildAll"
BuildingNow = True
Call Form_Load
BuildTimeElapsed = 0
HappyAdd = HappyAdd * BuildNum
SickRisk = SickRisk * BuildNum
LogsNeed = LogsNeed * BuildNum
BuildersNeed = BuildersNeed * BuildNum
End Sub

Private Sub cmdCancel_Click()
Unload Form2
Form1.tmrTime.Enabled = True
End Sub

Private Sub Form_Load()
Form1.tmrTime.Enabled = False
If BuildingNow = True Then
lblBuilding.Caption = "Currently Building " & BuildType
cmdCancel.Caption = "&Exit"
cmdBuild1.Enabled = False
cmdBuildAll.Enabled = False
lblDescription.Caption = ""
lstBuilding.Enabled = False
txtBuildNum.Enabled = False
Else
lblBuilding.Caption = "Choose a Building"
lstBuilding.Enabled = True
cmdCancel.Caption = "&Cancel"
End If
End Sub

Private Sub lstBuilding_Click()
'MsgBox ("selected is " & lstBuilding.Selected)
'MsgBox ("listindex is " & lstBuilding.ListIndex)
'MsgBox ("item is " & lstBuilding.Text)
'MsgBox ("selected is " & lstBuilding.Selected)
Select Case lstBuilding.ListIndex
    Case Is = 0: 'simple grass shelter
    BuildersNeed = 1
    BuildType = "Simple Grass Shelter"
    HappyAdd = 1
    FireRisk = 5
    SickRiskChange = -2
    LogsNeed = 50
    BuildTime = 1
    Case Is = 1: 'wigwam
    BuildersNeed = 2
    HappyAdd = 2
    FireRisk = 8
    SickRiskChange = -3
    BuildType = "Wigwam"
    LogsNeed = 60
    BuildTime = 2
    Case Is = 2: 'underground house
    BuildersNeed = 4
    HappyAdd = 1
    FireRisk = 2
    SickRiskChange = -5
    BuildType = "Underground House"
    LogsNeed = 55
    BuildTime = 7
    Case Is = 3: '4 room house
    BuildersNeed = 7
    HappyAdd = 3
    FireRisk = 4
    SickRiskChange = -20
    BuildType = "4 Room House"
    BuildTime = 10
    LogsNeed = 85
    Case Is = 4: '5 room house
    BuildersNeed = 10
    HappyAdd = 4
    FireRisk = 4
    SickRiskChange = -20
    BuildType = "5 Room House"
    BuildTime = 15
    LogsNeed = 100
End Select
lblDescription.Caption = "Happiness: +" & HappyAdd & vbCr & "Fire Risk: " & FireRisk & "%" & vbCr & "Sickness Risk: " & SickRiskChange & vbCr & "Wood Required: " & LogsNeed & vbCr & "Days to Build: " & BuildTime & vbCr & "Builders required: " & BuildersNeed
lblBuilding.Caption = BuildType
txtBuildNum.Enabled = True
'lblDescription.Caption = "This is line one" & vbCr & "This is line 2"
End Sub

Private Sub txtBuildNum_Change()
BuildNum = Val(txtBuildNum.Text)
If BuildersNeed <= Available And BuildNum <> 0 And BuildNum <> "" And Logs >= LogsNeed Then
cmdBuild1.Enabled = True
Else
cmdBuild1.Enabled = False
End If
If BuildNum * BuildersNeed <= Available And BuildNum <> 0 And BuildNum <> "" And Logs >= LogsNeed * BuildNum Then
cmdBuildAll.Enabled = True
Else
cmdBuildAll.Enabled = False
End If
End Sub
