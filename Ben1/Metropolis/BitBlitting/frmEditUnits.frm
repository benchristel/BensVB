VERSION 5.00
Begin VB.Form frmEditUnits 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Editor"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFire 
      Height          =   315
      Left            =   5400
      TabIndex        =   46
      Text            =   "2"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtHeal 
      Height          =   315
      Left            =   5400
      TabIndex        =   43
      Text            =   "2"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtThreshold 
      Height          =   315
      Left            =   5400
      TabIndex        =   42
      Text            =   "500"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   6420
      TabIndex        =   41
      Top             =   2820
      Width           =   855
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   435
      Left            =   7320
      TabIndex        =   40
      Top             =   2820
      Width           =   795
   End
   Begin VB.TextBox txtSpeed 
      Height          =   315
      Left            =   5400
      TabIndex        =   38
      Text            =   "35"
      Top             =   1440
      Width           =   975
   End
   Begin VB.ListBox lstSelect 
      Height          =   2400
      ItemData        =   "frmEditUnits.frx":0000
      Left            =   6420
      List            =   "frmEditUnits.frx":0043
      Sorted          =   -1  'True
      TabIndex        =   37
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtChicken 
      Height          =   315
      Left            =   5400
      TabIndex        =   33
      Text            =   "360"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtHP 
      Height          =   315
      Left            =   5400
      TabIndex        =   32
      Text            =   "600"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtHealth 
      Height          =   315
      Left            =   5400
      TabIndex        =   31
      Text            =   "600"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtRange 
      Height          =   315
      Index           =   6
      Left            =   2220
      TabIndex        =   27
      Text            =   "0"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtRange 
      Height          =   315
      Index           =   5
      Left            =   2220
      TabIndex        =   26
      Text            =   "30"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtRange 
      Height          =   315
      Index           =   4
      Left            =   2220
      TabIndex        =   25
      Text            =   "30"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtRange 
      Height          =   315
      Index           =   3
      Left            =   2220
      TabIndex        =   24
      Text            =   "0"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtRange 
      Height          =   315
      Index           =   2
      Left            =   2220
      TabIndex        =   23
      Text            =   "32"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtRange 
      Height          =   315
      Index           =   1
      Left            =   2220
      TabIndex        =   22
      Text            =   "40"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtRange 
      Height          =   315
      Index           =   0
      Left            =   2220
      TabIndex        =   21
      Text            =   "32"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtArmor 
      Height          =   315
      Index           =   6
      Left            =   3240
      TabIndex        =   20
      Text            =   "0"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtArmor 
      Height          =   315
      Index           =   5
      Left            =   3240
      TabIndex        =   19
      Text            =   "5"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtArmor 
      Height          =   315
      Index           =   4
      Left            =   3240
      TabIndex        =   18
      Text            =   "5"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtArmor 
      Height          =   315
      Index           =   3
      Left            =   3240
      TabIndex        =   17
      Text            =   "10"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtArmor 
      Height          =   315
      Index           =   2
      Left            =   3240
      TabIndex        =   16
      Text            =   "10"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtArmor 
      Height          =   315
      Index           =   1
      Left            =   3240
      TabIndex        =   15
      Text            =   "10"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtArmor 
      Height          =   315
      Index           =   0
      Left            =   3240
      TabIndex        =   14
      Text            =   "10"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtDamage 
      Height          =   315
      Index           =   6
      Left            =   1200
      TabIndex        =   13
      Text            =   "0"
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtDamage 
      Height          =   315
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Text            =   "3"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtDamage 
      Height          =   315
      Index           =   4
      Left            =   1200
      TabIndex        =   4
      Text            =   "3"
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtDamage 
      Height          =   315
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Text            =   "0"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtDamage 
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Text            =   "1"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtDamage 
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Text            =   "3"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtDamage 
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Text            =   "2"
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Of Fire"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   13
      Left            =   4260
      TabIndex        =   47
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Healing"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   12
      Left            =   4260
      TabIndex        =   45
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Threshold"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   4260
      TabIndex        =   44
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   4260
      TabIndex        =   39
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Chicken"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   4260
      TabIndex        =   36
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Start HP"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   4260
      TabIndex        =   35
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Health"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   4260
      TabIndex        =   34
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblColumn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Armor"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   30
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblColumn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Range"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   2220
      TabIndex        =   29
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblColumn 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Firepower"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   28
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Burn"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   60
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Crush"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   60
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Impact"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Slash"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Chop"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pierce"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   7
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblRow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Puncture"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
Dim i
For i = 0 To 6
PFirepower(i) = Val(txtDamage(i).Text)
Next i
For i = 0 To 6
PRange(i) = Val(txtRange(i).Text)
Next i
For i = 0 To 6
PArmor(i) = Val(txtArmor(i).Text)
Next i
PHealth = Val(txtHealth.Text)
PHP = Val(txtHP.Text)
PChicken = Val(txtChicken.Text)
PSpeed = Val(txtSpeed.Text)
PThreshold = Val(txtThreshold.Text)
PHealing = Val(txtHeal.Text)
PRateOfFire = Val(txtFire.Text)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call SetValues(PFirepower(0), PFirepower(1), PFirepower(2), PFirepower(3), PFirepower(4), PFirepower(5), PFirepower(6), _
    PRange(0), PRange(1), PRange(2), PRange(3), PRange(4), PRange(5), PRange(6), _
    PArmor(0), PArmor(1), PArmor(2), PArmor(3), PArmor(4), PArmor(5), PArmor(6), PHealth, PHP, PChicken, PSpeed, PThreshold, PHealing, PRateOfFire)
End Sub

Private Sub lstSelect_Click()
PName = lstSelect.Text
Select Case lstSelect.Text
'
'Puncture
'Pierce
'Chop
'Slash
'Crush
'Impact
'Burn
'

Case Is = "Archer"
Call SetValues(5, 3, 0, 0, 1, 1, 0, 860, 1550, 0, 0, 240, 240, 0, 20, 25, 15, 15, 15, 15, 8, 1100, 1100, 500, 38, 1000, 2, 3)
Case Is = "Artillery"
Call SetValues(15, 8, 0, 0, 5, 2, 2, 1500, 3000, 0, 0, 480, 600, 375, 50, 50, 65, 80, 60, 60, 95, 850, 850, 140, 20, 800, 1, 7)
Case Is = "Assassin"
Call SetValues(0, 0, 0, 0, 4, 3, 3, 0, 0, 0, 0, 360, 380, 440, 70, 70, 70, 70, 70, 70, 70, 600, 600, 100, 42, 550, 2, 2)
Case Is = "Clone Warrior"
Call SetValues(4, 3, 4, 3, 5, 4, 2, 460, 480, 450, 480, 420, 400, 300, 35, 35, 35, 35, 50, 50, 20, 450, 450, 0, 32, 450, 2, 2)
Case Is = "Crossbowman"
Call SetValues(8, 6, 0, 0, 2, 1, 0, 1200, 1560, 0, 0, 220, 240, 0, 25, 30, 25, 25, 25, 20, 10, 1200, 1200, 490, 32, 1100, 2, 4)
Case Is = "Horseman"
Call SetValues(4, 3, 1, 1, 0, 0, 0, 440, 500, 400, 400, 0, 0, 0, 15, 15, 20, 20, 70, 70, 4, 865, 865, 380, 56, 800, 1, 3)
Case Is = "Infantry"
Call SetValues(12, 10, 0, 0, 1, 1, 1, 1515, 2080, 0, 0, 400, 400, 480, 40, 40, 80, 80, 30, 30, 12, 1500, 1500, 450, 38, 1350, 3, 8)
Case Is = "Knight"
Call SetValues(5, 3, 4, 3, 1, 1, 0, 575, 615, 480, 460, 300, 300, 0, 60, 60, 80, 90, 50, 50, 40, 1000, 1000, 320, 50, 900, 2, 3)
Case Is = "Legion"
Call SetValues(5, 5, 6, 6, 0, 1, 0, 400, 480, 500, 500, 0, 360, 0, 30, 20, 30, 20, 15, 15, 8, 1000, 1000, 315, 30, 900, 2, 2)
Case Is = "Longbowman"
Call SetValues(7, 6, 0, 0, 1, 1, 0, 1580, 1800, 0, 0, 300, 300, 0, 30, 30, 30, 25, 25, 25, 8, 1250, 1250, 500, 32, 1100, 2, 6)
Case Is = "Mech"
Call SetValues(12, 10, 0, 0, 10, 6, 5, 2000, 2680, 0, 0, 300, 480, 960, 20, 20, 90, 90, 90, 70, 70, 2000, 2000, 800, 40, 1750, 3, 12)
Case Is = "Militia"
Call SetValues(2, 3, 1, 0, 3, 3, 0, 480, 600, 480, 0, 450, 450, 0, 10, 10, 10, 10, 5, 5, 0, 600, 600, 360, 35, 500, 2, 2)
End Select
End Sub

Private Sub SetValues(F1, F2, F3, F4, F5, F6, F7, R1, R2, R3, R4, R5, R6, R7, _
    A1, A2, A3, A4, A5, A6, A7, Health, HP, Chicken, Speed, Threshold, Healing, RateOfFire)
txtDamage(0).Text = F1
txtDamage(1).Text = F2
txtDamage(2).Text = F3
txtDamage(3).Text = F4
txtDamage(4).Text = F5
txtDamage(5).Text = F6
txtDamage(6).Text = F7
txtRange(0).Text = R1
txtRange(1).Text = R2
txtRange(2).Text = R3
txtRange(3).Text = R4
txtRange(4).Text = R5
txtRange(5).Text = R6
txtRange(6).Text = R7
txtArmor(0).Text = A1
txtArmor(1).Text = A2
txtArmor(2).Text = A3
txtArmor(3).Text = A4
txtArmor(4).Text = A5
txtArmor(5).Text = A6
txtArmor(6).Text = A7
txtHealth.Text = Health
txtHP.Text = HP
txtChicken.Text = Chicken
txtSpeed.Text = Speed
txtThreshold.Text = Threshold
txtHeal.Text = Healing
txtFire.Text = RateOfFire
End Sub

Private Sub Text1_Change()

End Sub
