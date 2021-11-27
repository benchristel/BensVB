VERSION 5.00
Begin VB.Form frmLoadout 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   5490
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSwitchPlayer 
      Caption         =   "Switch Player"
      Height          =   435
      Left            =   7800
      TabIndex        =   4
      Top             =   540
      Width           =   1155
   End
   Begin VB.CommandButton cmdMakeUnit 
      Caption         =   "Make Unit"
      Height          =   435
      Left            =   7800
      TabIndex        =   3
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdRename 
      Caption         =   "Rename"
      Height          =   435
      Left            =   6120
      TabIndex        =   1
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label lblPoints 
      Alignment       =   1  'Right Justify
      Caption         =   "2000"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6120
      TabIndex        =   5
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label lblSpendPoints 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label lblName 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   60
      Width           =   2115
   End
   Begin VB.Image imgHardpoint 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   3
      Left            =   5520
      Top             =   480
      Width           =   540
   End
   Begin VB.Image imgHardpoint 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   2
      Left            =   4980
      Top             =   480
      Width           =   540
   End
   Begin VB.Image imgHardpoint 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   1
      Left            =   4440
      Top             =   480
      Width           =   540
   End
   Begin VB.Image imgHardpoint 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Index           =   0
      Left            =   3900
      Top             =   480
      Width           =   540
   End
   Begin VB.Image imgUnitType 
      Height          =   900
      Left            =   2940
      Picture         =   "frmLoadout.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   900
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   11
      Left            =   1560
      Picture         =   "frmLoadout.frx":2A74
      Top             =   1620
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   10
      Left            =   1080
      Picture         =   "frmLoadout.frx":373E
      Top             =   1620
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   9
      Left            =   600
      Picture         =   "frmLoadout.frx":4408
      Top             =   1620
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   8
      Left            =   120
      Picture         =   "frmLoadout.frx":4CD2
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   7
      Left            =   1080
      Picture         =   "frmLoadout.frx":559C
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   6
      Left            =   1560
      Picture         =   "frmLoadout.frx":5E66
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   5
      Left            =   1560
      Picture         =   "frmLoadout.frx":6730
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   4
      Left            =   120
      Picture         =   "frmLoadout.frx":6FFA
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   3
      Left            =   600
      Picture         =   "frmLoadout.frx":78C4
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   2
      Left            =   1080
      Picture         =   "frmLoadout.frx":858E
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   1
      Left            =   600
      Picture         =   "frmLoadout.frx":8E58
      Top             =   540
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmLoadout.frx":9B22
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   1995
   End
End
Attribute VB_Name = "frmLoadout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PlaceWeapon, Hardpoint(0 To 3), ActivePlayer
Dim MakeX, MakeY, UnitType
Dim Points(1 To 2) As Integer, SpendPoints As Integer
Private Sub cmdMakeUnit_Click()
Dim i
cmdSwitchPlayer.Enabled = True
If SpendPoints <= Points(ActivePlayer) Then
If ActivePlayer = 1 Then MakeY = 100
If ActivePlayer = 2 Then MakeY = 1200
MakeX = MakeX + 100
lblName.Caption = GenerateName
Call LoadUnit(UnitType, MakeX, MakeY, ActivePlayer, Hardpoint(0), Hardpoint(1), Hardpoint(2), Hardpoint(3), lblName.Caption)
Points(ActivePlayer) = Points(ActivePlayer) - SpendPoints
lblPoints.Caption = Points(ActivePlayer)
'For i = 0 To 3
'Hardpoint(i) = UData(UnitType).Weapon(i)
'imgHardpoint(i).Picture = Nothing
'Next i
SpendPoints = UData(UnitType).Cost
Call CalculatePoints
End If
End Sub

Private Sub cmdRename_Click()
Dim HoldName
HoldName = lblName.Caption
lblName.Caption = InputBox("Type a new name:", "Rename Unit", GenerateName)
If lblName.Caption = "" Then lblName.Caption = HoldName
End Sub

Private Sub cmdSwitchPlayer_Click()
cmdSwitchPlayer.Enabled = False
ActivePlayer = ActivePlayer + 1
If ActivePlayer > Teams Then
Load frmInterface
frmInterface.Visible = True
Unload Me
Exit Sub
End If
MakeX = 0
lblPoints.Caption = Points(ActivePlayer)
End Sub

Private Sub Form_Load()
Teams = 2
cmdSwitchPlayer.Enabled = False
ActivePlayer = 1
Call InitiateUnits
UnitType = 1
Points(1) = 3000
Points(2) = 3000
SpendPoints = UData(UnitType).Cost
lblSpendPoints.Caption = SpendPoints
lblPoints.Caption = Points(1)
lblName.Caption = GenerateName
End Sub

Private Sub imgHardpoint_Click(Index As Integer)
Dim i
For i = 0 To 3
If Hardpoint(i) = PlaceWeapon And PlaceWeapon <> 0 Then Exit Sub
Next i
If Hardpoint(Index) > -1 Then
SpendPoints = SpendPoints - Weapon(Hardpoint(Index)).Cost
End If
    If PlaceWeapon = 0 Then 'empty slot
    Hardpoint(Index) = -1
    imgHardpoint(Index).Picture = Nothing
    Else
    Hardpoint(Index) = PlaceWeapon
    imgHardpoint(Index).Picture = imgWeapon(PlaceWeapon).Picture
    End If
    Call CalculatePoints
End Sub

Private Sub imgWeapon_Click(Index As Integer)
    PlaceWeapon = Index
    frmLoadout.MouseIcon = imgWeapon(Index).Picture
End Sub

Private Sub CalculatePoints()
SpendPoints = UData(UnitType).Cost
    Dim AWeapon(0 To 3), i
    For i = 0 To 3
    AWeapon(i) = Hardpoint(i)
    If AWeapon(i) = -1 Then AWeapon(i) = 0
    SpendPoints = SpendPoints + Weapon(AWeapon(i)).Cost
    Next i
    lblSpendPoints.Caption = SpendPoints
End Sub

