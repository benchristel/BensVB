VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00004080&
   BorderStyle     =   0  'None
   ClientHeight    =   11670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15240
   ControlBox      =   0   'False
   Icon            =   "frmLuna.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   11670
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdEscape 
      Caption         =   "Escape"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      MousePointer    =   1  'Arrow
      TabIndex        =   9
      Top             =   11040
      Width           =   915
   End
   Begin VB.Image imgPyromegalon 
      Height          =   720
      Left            =   14400
      Picture         =   "frmLuna.frx":08CA
      Top             =   6780
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgCryomegalon 
      Height          =   720
      Left            =   14400
      Picture         =   "frmLuna.frx":2594
      Top             =   6720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuildUnit 
      Height          =   720
      Index           =   6
      Left            =   9420
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":425E
      ToolTipText     =   "Cryomegalon"
      Top             =   10260
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuildUnit 
      Height          =   720
      Index           =   5
      Left            =   9480
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":5F28
      ToolTipText     =   "Pyromegalon"
      Top             =   10260
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuilding 
      Height          =   720
      Index           =   8
      Left            =   6060
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":7BF2
      ToolTipText     =   "Cryogenics Lab"
      Top             =   10260
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgPlains 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":98BC
      Top             =   2880
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgGlacier 
      Height          =   720
      Left            =   14400
      Picture         =   "frmLuna.frx":B586
      Top             =   2880
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgGribbleBlue 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":D250
      Top             =   6540
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgGribbleRed 
      Height          =   720
      Left            =   14520
      Picture         =   "frmLuna.frx":EF1A
      Top             =   6540
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuildUnit 
      Height          =   720
      Index           =   4
      Left            =   8760
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":10BE4
      ToolTipText     =   "Gribble"
      Top             =   10260
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuilding 
      Height          =   720
      Index           =   7
      Left            =   7500
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":128AE
      ToolTipText     =   "Gribble Hatchery"
      Top             =   10260
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgHatcheryRed 
      Height          =   720
      Left            =   14400
      Picture         =   "frmLuna.frx":14578
      Top             =   9480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgHatcheryBlue 
      Height          =   720
      Left            =   14400
      Picture         =   "frmLuna.frx":16242
      Top             =   10200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuilding 
      Height          =   720
      Index           =   6
      Left            =   6060
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":17F0C
      ToolTipText     =   "Geothermal Plant"
      Top             =   10260
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgGeothermalRed 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":19BD6
      Top             =   8220
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgCryogenicsBlue 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":1B8A0
      Top             =   6780
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgLava 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":1D56A
      Top             =   3300
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgLifeSupportRed 
      Height          =   720
      Left            =   13680
      Picture         =   "frmLuna.frx":1F234
      Top             =   10200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgLifeSupportBlue 
      Height          =   720
      Left            =   13680
      Picture         =   "frmLuna.frx":20EFE
      Top             =   9480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgSupportIcon 
      Height          =   240
      Left            =   14520
      Picture         =   "frmLuna.frx":22BC8
      Top             =   960
      Width           =   240
   End
   Begin VB.Label lblSupport 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1/3"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   14460
      TabIndex        =   8
      Top             =   900
      Width           =   855
   End
   Begin VB.Image imgBuilding 
      Height          =   720
      Index           =   5
      Left            =   5340
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":22F52
      ToolTipText     =   "Life Support"
      Top             =   9540
      Width           =   720
   End
   Begin VB.Image imgMechLabRed 
      Height          =   720
      Left            =   12240
      Picture         =   "frmLuna.frx":24C1C
      Top             =   9480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgMechLabBlue 
      Height          =   720
      Left            =   12240
      Picture         =   "frmLuna.frx":268E6
      Top             =   10200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuilding 
      Height          =   720
      Index           =   4
      Left            =   7500
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":285B0
      ToolTipText     =   "Mech Lab"
      Top             =   9540
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgTankRed 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":2A27A
      Top             =   6420
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgTankBlue 
      Height          =   720
      Left            =   14520
      Picture         =   "frmLuna.frx":2BF44
      Top             =   6480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuildUnit 
      Height          =   720
      Index           =   3
      Left            =   10920
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":2DC0E
      ToolTipText     =   "Tank"
      Top             =   9540
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgLuna 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":2F8D8
      Top             =   2880
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgJediBlue 
      Height          =   720
      Left            =   14520
      Picture         =   "frmLuna.frx":315A2
      Top             =   6420
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgJediRed 
      Height          =   720
      Left            =   14520
      Picture         =   "frmLuna.frx":3326C
      Top             =   6480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuildUnit 
      Height          =   720
      Index           =   2
      Left            =   10200
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":34F36
      ToolTipText     =   "Jedi"
      Top             =   9660
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgDefenseTowerRed 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":36C00
      Top             =   7500
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgDefenseTowerBlue 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":388CA
      Top             =   6780
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuilding 
      Height          =   720
      Index           =   3
      Left            =   6780
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":3A594
      ToolTipText     =   "Defense Tower"
      Top             =   10260
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblViewBuilding 
      BackColor       =   &H00000000&
      Caption         =   "Rocky Outcropping"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   10440
      Width           =   3015
   End
   Begin VB.Label lblViewUnit 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   10620
      Width           =   3015
   End
   Begin VB.Label lblActiveUnit 
      BackColor       =   &H00000000&
      Caption         =   "Blue Worker HP 100/100 Moves: 1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   10800
      Width           =   3015
   End
   Begin VB.Image imgUnitMode 
      Height          =   480
      Index           =   3
      Left            =   3180
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":3C25E
      ToolTipText     =   "Attack"
      Top             =   10440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line linViewUnitHealthBar 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   1020
      X2              =   1740
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Image imgViewUnit 
      Height          =   720
      Left            =   1020
      Top             =   9600
      Width           =   720
   End
   Begin VB.Image imgTrooperRed 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":3CF28
      Top             =   6480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgTrooperBlue 
      Height          =   720
      Left            =   14220
      Picture         =   "frmLuna.frx":3EBF2
      Top             =   6720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuildUnit 
      Height          =   720
      Index           =   1
      Left            =   9480
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":408BC
      ToolTipText     =   "Trooper"
      Top             =   9540
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuilding 
      Height          =   720
      Index           =   2
      Left            =   6780
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":42586
      ToolTipText     =   "Barracks"
      Top             =   9540
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgUnitMode 
      Height          =   480
      Index           =   2
      Left            =   3720
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":44250
      ToolTipText     =   "Skip"
      Top             =   9900
      Width           =   480
   End
   Begin VB.Image imgBarracksRed 
      Height          =   720
      Left            =   12960
      Picture         =   "frmLuna.frx":44F1A
      Top             =   10200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBarracksBlue 
      Height          =   720
      Left            =   12960
      Picture         =   "frmLuna.frx":46BE4
      Top             =   9480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgViewBuilding 
      Height          =   720
      Left            =   120
      Picture         =   "frmLuna.frx":488AE
      Top             =   9600
      Width           =   720
   End
   Begin VB.Line linBldgHealthBar 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   120
      X2              =   840
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Image imgHighlight 
      Height          =   720
      Left            =   0
      Picture         =   "frmLuna.frx":4A578
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgBuildUnit 
      Height          =   720
      Index           =   0
      Left            =   8760
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":4C242
      ToolTipText     =   "Worker"
      Top             =   9540
      Width           =   720
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   8760
      MousePointer    =   1  'Arrow
      TabIndex        =   4
      Top             =   9600
      Width           =   2895
   End
   Begin VB.Image imgConstruct2Red 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":4DF0C
      Top             =   5100
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgConstruct1Red 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":4FBD6
      Top             =   4320
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgConstruct2Blue 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":518A0
      Top             =   4380
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgConstruct1Blue 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":5356A
      Top             =   3660
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Line linHealthBar 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   1920
      X2              =   2640
      Y1              =   9480
      Y2              =   9480
   End
   Begin VB.Image imgActiveUnit 
      Height          =   720
      Left            =   1920
      Picture         =   "frmLuna.frx":55234
      Top             =   9600
      Width           =   720
   End
   Begin VB.Image imgVespianMineRed 
      Height          =   720
      Left            =   10140
      Picture         =   "frmLuna.frx":56EFE
      Top             =   9540
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgVespianMineBlue 
      Height          =   720
      Left            =   12180
      Picture         =   "frmLuna.frx":58BC8
      Top             =   9480
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgOreMineRed 
      Height          =   720
      Left            =   10140
      Picture         =   "frmLuna.frx":5A892
      Top             =   10260
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgOreMineBlue 
      Height          =   720
      Left            =   12480
      Picture         =   "frmLuna.frx":5C55C
      Top             =   10200
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBuilding 
      Height          =   720
      Index           =   1
      Left            =   6060
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":5E226
      ToolTipText     =   "Vespian Mine"
      Top             =   9540
      Width           =   720
   End
   Begin VB.Image imgBuilding 
      Height          =   720
      Index           =   0
      Left            =   5340
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":5FEF0
      ToolTipText     =   "Ore Mine"
      Top             =   10260
      Width           =   720
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   5340
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      Top             =   9540
      Width           =   2895
   End
   Begin VB.Image imgVespianIcon 
      Height          =   240
      Left            =   14520
      Picture         =   "frmLuna.frx":61BBA
      Top             =   540
      Width           =   240
   End
   Begin VB.Image imgOreIcon 
      Height          =   240
      Left            =   14520
      Picture         =   "frmLuna.frx":61F44
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblVespian 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   14460
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblOre 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   14460
      TabIndex        =   0
      Top             =   60
      Width           =   855
   End
   Begin VB.Image imgWorkerRed 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":622CE
      Top             =   6000
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgWorkerBlue 
      Height          =   720
      Left            =   14400
      Picture         =   "frmLuna.frx":63F98
      Top             =   5940
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgUnit 
      Height          =   735
      Index           =   0
      Left            =   14460
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgRedTownHall 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":65C62
      Top             =   2160
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgBlueTownHall 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":6792C
      Top             =   1440
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgOre 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":695F6
      Top             =   720
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgVespian 
      Height          =   720
      Left            =   14460
      Picture         =   "frmLuna.frx":6B2C0
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imgUnitMode 
      Height          =   480
      Index           =   1
      Left            =   3180
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":6CF8A
      ToolTipText     =   "Build"
      Top             =   10440
      Width           =   480
   End
   Begin VB.Image imgUnitMode 
      Height          =   480
      Index           =   0
      Left            =   3180
      MousePointer    =   1  'Arrow
      Picture         =   "frmLuna.frx":6DC54
      ToolTipText     =   "Move"
      Top             =   9900
      Width           =   480
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   259
      Left            =   13680
      Picture         =   "frmLuna.frx":6E91E
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   258
      Left            =   12960
      Picture         =   "frmLuna.frx":705E8
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   257
      Left            =   12240
      Picture         =   "frmLuna.frx":722B2
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   256
      Left            =   11520
      Picture         =   "frmLuna.frx":73F7C
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   255
      Left            =   10800
      Picture         =   "frmLuna.frx":75C46
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   254
      Left            =   10080
      Picture         =   "frmLuna.frx":77910
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   253
      Left            =   9360
      Picture         =   "frmLuna.frx":795DA
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   252
      Left            =   8640
      Picture         =   "frmLuna.frx":7B2A4
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   251
      Left            =   7920
      Picture         =   "frmLuna.frx":7CF6E
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   250
      Left            =   7200
      Picture         =   "frmLuna.frx":7EC38
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   249
      Left            =   6480
      Picture         =   "frmLuna.frx":80902
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   248
      Left            =   5760
      Picture         =   "frmLuna.frx":825CC
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   247
      Left            =   5040
      Picture         =   "frmLuna.frx":84296
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   246
      Left            =   4320
      Picture         =   "frmLuna.frx":85F60
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   245
      Left            =   3600
      Picture         =   "frmLuna.frx":87C2A
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   244
      Left            =   2880
      Picture         =   "frmLuna.frx":898F4
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   243
      Left            =   2160
      Picture         =   "frmLuna.frx":8B5BE
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   242
      Left            =   1440
      Picture         =   "frmLuna.frx":8D288
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   241
      Left            =   720
      Picture         =   "frmLuna.frx":8EF52
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   240
      Left            =   0
      Picture         =   "frmLuna.frx":90C1C
      Top             =   8640
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   239
      Left            =   13680
      Picture         =   "frmLuna.frx":928E6
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   238
      Left            =   12960
      Picture         =   "frmLuna.frx":945B0
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   237
      Left            =   12240
      Picture         =   "frmLuna.frx":9627A
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   236
      Left            =   11520
      Picture         =   "frmLuna.frx":97F44
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   235
      Left            =   10800
      Picture         =   "frmLuna.frx":99C0E
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   234
      Left            =   10080
      Picture         =   "frmLuna.frx":9B8D8
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   233
      Left            =   9360
      Picture         =   "frmLuna.frx":9D5A2
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   232
      Left            =   8640
      Picture         =   "frmLuna.frx":9F26C
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   231
      Left            =   7920
      Picture         =   "frmLuna.frx":A0F36
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   230
      Left            =   7200
      Picture         =   "frmLuna.frx":A2C00
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   229
      Left            =   6480
      Picture         =   "frmLuna.frx":A48CA
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   228
      Left            =   5760
      Picture         =   "frmLuna.frx":A6594
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   227
      Left            =   5040
      Picture         =   "frmLuna.frx":A825E
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   226
      Left            =   4320
      Picture         =   "frmLuna.frx":A9F28
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   225
      Left            =   3600
      Picture         =   "frmLuna.frx":ABBF2
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   224
      Left            =   2880
      Picture         =   "frmLuna.frx":AD8BC
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   223
      Left            =   2160
      Picture         =   "frmLuna.frx":AF586
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   222
      Left            =   1440
      Picture         =   "frmLuna.frx":B1250
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   221
      Left            =   720
      Picture         =   "frmLuna.frx":B2F1A
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   220
      Left            =   0
      Picture         =   "frmLuna.frx":B4BE4
      Top             =   7920
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   219
      Left            =   13680
      Picture         =   "frmLuna.frx":B68AE
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   218
      Left            =   12960
      Picture         =   "frmLuna.frx":B8578
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   217
      Left            =   12240
      Picture         =   "frmLuna.frx":BA242
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   216
      Left            =   11520
      Picture         =   "frmLuna.frx":BBF0C
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   215
      Left            =   10800
      Picture         =   "frmLuna.frx":BDBD6
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   214
      Left            =   10080
      Picture         =   "frmLuna.frx":BF8A0
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   213
      Left            =   9360
      Picture         =   "frmLuna.frx":C156A
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   212
      Left            =   8640
      Picture         =   "frmLuna.frx":C3234
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   211
      Left            =   7920
      Picture         =   "frmLuna.frx":C4EFE
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   210
      Left            =   7200
      Picture         =   "frmLuna.frx":C6BC8
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   209
      Left            =   6480
      Picture         =   "frmLuna.frx":C8892
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   208
      Left            =   5760
      Picture         =   "frmLuna.frx":CA55C
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   207
      Left            =   5040
      Picture         =   "frmLuna.frx":CC226
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   206
      Left            =   4320
      Picture         =   "frmLuna.frx":CDEF0
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   205
      Left            =   3600
      Picture         =   "frmLuna.frx":CFBBA
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   204
      Left            =   2880
      Picture         =   "frmLuna.frx":D1884
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   203
      Left            =   2160
      Picture         =   "frmLuna.frx":D354E
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   202
      Left            =   1440
      Picture         =   "frmLuna.frx":D5218
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   201
      Left            =   720
      Picture         =   "frmLuna.frx":D6EE2
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   200
      Left            =   0
      Picture         =   "frmLuna.frx":D8BAC
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   199
      Left            =   13680
      Picture         =   "frmLuna.frx":DA876
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   198
      Left            =   12960
      Picture         =   "frmLuna.frx":DC540
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   197
      Left            =   12240
      Picture         =   "frmLuna.frx":DE20A
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   196
      Left            =   11520
      Picture         =   "frmLuna.frx":DFED4
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   195
      Left            =   10800
      Picture         =   "frmLuna.frx":E1B9E
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   194
      Left            =   10080
      Picture         =   "frmLuna.frx":E3868
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   193
      Left            =   9360
      Picture         =   "frmLuna.frx":E5532
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   192
      Left            =   8640
      Picture         =   "frmLuna.frx":E71FC
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   191
      Left            =   7920
      Picture         =   "frmLuna.frx":E8EC6
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   190
      Left            =   7200
      Picture         =   "frmLuna.frx":EAB90
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   189
      Left            =   6480
      Picture         =   "frmLuna.frx":EC85A
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   188
      Left            =   5760
      Picture         =   "frmLuna.frx":EE524
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   187
      Left            =   5040
      Picture         =   "frmLuna.frx":F01EE
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   186
      Left            =   4320
      Picture         =   "frmLuna.frx":F1EB8
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   185
      Left            =   3600
      Picture         =   "frmLuna.frx":F3B82
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   184
      Left            =   2880
      Picture         =   "frmLuna.frx":F584C
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   183
      Left            =   2160
      Picture         =   "frmLuna.frx":F7516
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   182
      Left            =   1440
      Picture         =   "frmLuna.frx":F91E0
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   181
      Left            =   720
      Picture         =   "frmLuna.frx":FAEAA
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   180
      Left            =   0
      Picture         =   "frmLuna.frx":FCB74
      Top             =   6480
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   179
      Left            =   13680
      Picture         =   "frmLuna.frx":FE83E
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   178
      Left            =   12960
      Picture         =   "frmLuna.frx":100508
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   177
      Left            =   12240
      Picture         =   "frmLuna.frx":1021D2
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   176
      Left            =   11520
      Picture         =   "frmLuna.frx":103E9C
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   175
      Left            =   10800
      Picture         =   "frmLuna.frx":105B66
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   174
      Left            =   10080
      Picture         =   "frmLuna.frx":107830
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   173
      Left            =   9360
      Picture         =   "frmLuna.frx":1094FA
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   172
      Left            =   8640
      Picture         =   "frmLuna.frx":10B1C4
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   171
      Left            =   7920
      Picture         =   "frmLuna.frx":10CE8E
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   170
      Left            =   7200
      Picture         =   "frmLuna.frx":10EB58
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   169
      Left            =   6480
      Picture         =   "frmLuna.frx":110822
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   168
      Left            =   5760
      Picture         =   "frmLuna.frx":1124EC
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   167
      Left            =   5040
      Picture         =   "frmLuna.frx":1141B6
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   166
      Left            =   4320
      Picture         =   "frmLuna.frx":115E80
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   165
      Left            =   3600
      Picture         =   "frmLuna.frx":117B4A
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   164
      Left            =   2880
      Picture         =   "frmLuna.frx":119814
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   163
      Left            =   2160
      Picture         =   "frmLuna.frx":11B4DE
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   162
      Left            =   1440
      Picture         =   "frmLuna.frx":11D1A8
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   161
      Left            =   720
      Picture         =   "frmLuna.frx":11EE72
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   160
      Left            =   0
      Picture         =   "frmLuna.frx":120B3C
      Top             =   5760
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   159
      Left            =   13680
      Picture         =   "frmLuna.frx":122806
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   158
      Left            =   12960
      Picture         =   "frmLuna.frx":1244D0
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   157
      Left            =   12240
      Picture         =   "frmLuna.frx":12619A
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   156
      Left            =   11520
      Picture         =   "frmLuna.frx":127E64
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   155
      Left            =   10800
      Picture         =   "frmLuna.frx":129B2E
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   154
      Left            =   10080
      Picture         =   "frmLuna.frx":12B7F8
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   153
      Left            =   9360
      Picture         =   "frmLuna.frx":12D4C2
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   152
      Left            =   8640
      Picture         =   "frmLuna.frx":12F18C
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   151
      Left            =   7920
      Picture         =   "frmLuna.frx":130E56
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   150
      Left            =   7200
      Picture         =   "frmLuna.frx":132B20
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   149
      Left            =   6480
      Picture         =   "frmLuna.frx":1347EA
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   148
      Left            =   5760
      Picture         =   "frmLuna.frx":1364B4
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   147
      Left            =   5040
      Picture         =   "frmLuna.frx":13817E
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   146
      Left            =   4320
      Picture         =   "frmLuna.frx":139E48
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   145
      Left            =   3600
      Picture         =   "frmLuna.frx":13BB12
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   144
      Left            =   2880
      Picture         =   "frmLuna.frx":13D7DC
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   143
      Left            =   2160
      Picture         =   "frmLuna.frx":13F4A6
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   142
      Left            =   1440
      Picture         =   "frmLuna.frx":141170
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   141
      Left            =   720
      Picture         =   "frmLuna.frx":142E3A
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   140
      Left            =   0
      Picture         =   "frmLuna.frx":144B04
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   139
      Left            =   13680
      Picture         =   "frmLuna.frx":1467CE
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   138
      Left            =   12960
      Picture         =   "frmLuna.frx":148498
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   137
      Left            =   12240
      Picture         =   "frmLuna.frx":14A162
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   136
      Left            =   11520
      Picture         =   "frmLuna.frx":14BE2C
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   135
      Left            =   10800
      Picture         =   "frmLuna.frx":14DAF6
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   134
      Left            =   10080
      Picture         =   "frmLuna.frx":14F7C0
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   133
      Left            =   9360
      Picture         =   "frmLuna.frx":15148A
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   132
      Left            =   8640
      Picture         =   "frmLuna.frx":153154
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   131
      Left            =   7920
      Picture         =   "frmLuna.frx":154E1E
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   130
      Left            =   7200
      Picture         =   "frmLuna.frx":156AE8
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   129
      Left            =   6480
      Picture         =   "frmLuna.frx":1587B2
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   128
      Left            =   5760
      Picture         =   "frmLuna.frx":15A47C
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   127
      Left            =   5040
      Picture         =   "frmLuna.frx":15C146
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   126
      Left            =   4320
      Picture         =   "frmLuna.frx":15DE10
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   125
      Left            =   3600
      Picture         =   "frmLuna.frx":15FADA
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   124
      Left            =   2880
      Picture         =   "frmLuna.frx":1617A4
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   123
      Left            =   2160
      Picture         =   "frmLuna.frx":16346E
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   122
      Left            =   1440
      Picture         =   "frmLuna.frx":165138
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   121
      Left            =   720
      Picture         =   "frmLuna.frx":166E02
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   120
      Left            =   0
      Picture         =   "frmLuna.frx":168ACC
      Top             =   4320
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   119
      Left            =   13680
      Picture         =   "frmLuna.frx":16A796
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   118
      Left            =   12960
      Picture         =   "frmLuna.frx":16C460
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   117
      Left            =   12240
      Picture         =   "frmLuna.frx":16E12A
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   116
      Left            =   11520
      Picture         =   "frmLuna.frx":16FDF4
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   115
      Left            =   10800
      Picture         =   "frmLuna.frx":171ABE
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   114
      Left            =   10080
      Picture         =   "frmLuna.frx":173788
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   113
      Left            =   9360
      Picture         =   "frmLuna.frx":175452
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   112
      Left            =   8640
      Picture         =   "frmLuna.frx":17711C
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   111
      Left            =   7920
      Picture         =   "frmLuna.frx":178DE6
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   110
      Left            =   7200
      Picture         =   "frmLuna.frx":17AAB0
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   109
      Left            =   6480
      Picture         =   "frmLuna.frx":17C77A
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   108
      Left            =   5760
      Picture         =   "frmLuna.frx":17E444
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   107
      Left            =   5040
      Picture         =   "frmLuna.frx":18010E
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   106
      Left            =   4320
      Picture         =   "frmLuna.frx":181DD8
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   105
      Left            =   3600
      Picture         =   "frmLuna.frx":183AA2
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   104
      Left            =   2880
      Picture         =   "frmLuna.frx":18576C
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   103
      Left            =   2160
      Picture         =   "frmLuna.frx":187436
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   102
      Left            =   1440
      Picture         =   "frmLuna.frx":189100
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   101
      Left            =   720
      Picture         =   "frmLuna.frx":18ADCA
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   100
      Left            =   0
      Picture         =   "frmLuna.frx":18CA94
      Top             =   3600
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   99
      Left            =   13680
      Picture         =   "frmLuna.frx":18E75E
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   98
      Left            =   12960
      Picture         =   "frmLuna.frx":190428
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   97
      Left            =   12240
      Picture         =   "frmLuna.frx":1920F2
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   96
      Left            =   11520
      Picture         =   "frmLuna.frx":193DBC
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   95
      Left            =   10800
      Picture         =   "frmLuna.frx":195A86
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   94
      Left            =   10080
      Picture         =   "frmLuna.frx":197750
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   93
      Left            =   9360
      Picture         =   "frmLuna.frx":19941A
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   92
      Left            =   8640
      Picture         =   "frmLuna.frx":19B0E4
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   91
      Left            =   7920
      Picture         =   "frmLuna.frx":19CDAE
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   90
      Left            =   7200
      Picture         =   "frmLuna.frx":19EA78
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   89
      Left            =   6480
      Picture         =   "frmLuna.frx":1A0742
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   88
      Left            =   5760
      Picture         =   "frmLuna.frx":1A240C
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   87
      Left            =   5040
      Picture         =   "frmLuna.frx":1A40D6
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   86
      Left            =   4320
      Picture         =   "frmLuna.frx":1A5DA0
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   85
      Left            =   3600
      Picture         =   "frmLuna.frx":1A7A6A
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   84
      Left            =   2880
      Picture         =   "frmLuna.frx":1A9734
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   83
      Left            =   2160
      Picture         =   "frmLuna.frx":1AB3FE
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   82
      Left            =   1440
      Picture         =   "frmLuna.frx":1AD0C8
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   81
      Left            =   720
      Picture         =   "frmLuna.frx":1AED92
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   80
      Left            =   0
      Picture         =   "frmLuna.frx":1B0A5C
      Top             =   2880
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   79
      Left            =   13680
      Picture         =   "frmLuna.frx":1B2726
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   78
      Left            =   12960
      Picture         =   "frmLuna.frx":1B43F0
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   77
      Left            =   12240
      Picture         =   "frmLuna.frx":1B60BA
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   76
      Left            =   11520
      Picture         =   "frmLuna.frx":1B7D84
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   75
      Left            =   10800
      Picture         =   "frmLuna.frx":1B9A4E
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   74
      Left            =   10080
      Picture         =   "frmLuna.frx":1BB718
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   73
      Left            =   9360
      Picture         =   "frmLuna.frx":1BD3E2
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   72
      Left            =   8640
      Picture         =   "frmLuna.frx":1BF0AC
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   71
      Left            =   7920
      Picture         =   "frmLuna.frx":1C0D76
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   70
      Left            =   7200
      Picture         =   "frmLuna.frx":1C2A40
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   69
      Left            =   6480
      Picture         =   "frmLuna.frx":1C470A
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   68
      Left            =   5760
      Picture         =   "frmLuna.frx":1C63D4
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   67
      Left            =   5040
      Picture         =   "frmLuna.frx":1C809E
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   66
      Left            =   4320
      Picture         =   "frmLuna.frx":1C9D68
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   65
      Left            =   3600
      Picture         =   "frmLuna.frx":1CBA32
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   64
      Left            =   2880
      Picture         =   "frmLuna.frx":1CD6FC
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   63
      Left            =   2160
      Picture         =   "frmLuna.frx":1CF3C6
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   62
      Left            =   1440
      Picture         =   "frmLuna.frx":1D1090
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   61
      Left            =   720
      Picture         =   "frmLuna.frx":1D2D5A
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   60
      Left            =   0
      Picture         =   "frmLuna.frx":1D4A24
      Top             =   2160
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   59
      Left            =   13680
      Picture         =   "frmLuna.frx":1D66EE
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   58
      Left            =   12960
      Picture         =   "frmLuna.frx":1D83B8
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   57
      Left            =   12240
      Picture         =   "frmLuna.frx":1DA082
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   56
      Left            =   11520
      Picture         =   "frmLuna.frx":1DBD4C
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   55
      Left            =   10800
      Picture         =   "frmLuna.frx":1DDA16
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   54
      Left            =   10080
      Picture         =   "frmLuna.frx":1DF6E0
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   53
      Left            =   9360
      Picture         =   "frmLuna.frx":1E13AA
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   52
      Left            =   8640
      Picture         =   "frmLuna.frx":1E3074
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   51
      Left            =   7920
      Picture         =   "frmLuna.frx":1E4D3E
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   50
      Left            =   7200
      Picture         =   "frmLuna.frx":1E6A08
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   49
      Left            =   6480
      Picture         =   "frmLuna.frx":1E86D2
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   48
      Left            =   5760
      Picture         =   "frmLuna.frx":1EA39C
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   47
      Left            =   5040
      Picture         =   "frmLuna.frx":1EC066
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   46
      Left            =   4320
      Picture         =   "frmLuna.frx":1EDD30
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   45
      Left            =   3600
      Picture         =   "frmLuna.frx":1EF9FA
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   44
      Left            =   2880
      Picture         =   "frmLuna.frx":1F16C4
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   43
      Left            =   2160
      Picture         =   "frmLuna.frx":1F338E
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   42
      Left            =   1440
      Picture         =   "frmLuna.frx":1F5058
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   41
      Left            =   720
      Picture         =   "frmLuna.frx":1F6D22
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   40
      Left            =   0
      Picture         =   "frmLuna.frx":1F89EC
      Top             =   1440
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   39
      Left            =   13680
      Picture         =   "frmLuna.frx":1FA6B6
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   38
      Left            =   12960
      Picture         =   "frmLuna.frx":1FC380
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   37
      Left            =   12240
      Picture         =   "frmLuna.frx":1FE04A
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   36
      Left            =   11520
      Picture         =   "frmLuna.frx":1FFD14
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   35
      Left            =   10800
      Picture         =   "frmLuna.frx":2019DE
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   34
      Left            =   10080
      Picture         =   "frmLuna.frx":2036A8
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   33
      Left            =   9360
      Picture         =   "frmLuna.frx":205372
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   32
      Left            =   8640
      Picture         =   "frmLuna.frx":20703C
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   31
      Left            =   7920
      Picture         =   "frmLuna.frx":208D06
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   30
      Left            =   7200
      Picture         =   "frmLuna.frx":20A9D0
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   29
      Left            =   6480
      Picture         =   "frmLuna.frx":20C69A
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   28
      Left            =   5760
      Picture         =   "frmLuna.frx":20E364
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   27
      Left            =   5040
      Picture         =   "frmLuna.frx":21002E
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   26
      Left            =   4320
      Picture         =   "frmLuna.frx":211CF8
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   25
      Left            =   3600
      Picture         =   "frmLuna.frx":2139C2
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   24
      Left            =   2880
      Picture         =   "frmLuna.frx":21568C
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   23
      Left            =   2160
      Picture         =   "frmLuna.frx":217356
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   22
      Left            =   1440
      Picture         =   "frmLuna.frx":219020
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   21
      Left            =   720
      Picture         =   "frmLuna.frx":21ACEA
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   20
      Left            =   0
      Picture         =   "frmLuna.frx":21C9B4
      Top             =   720
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   19
      Left            =   13680
      Picture         =   "frmLuna.frx":21E67E
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   18
      Left            =   12960
      Picture         =   "frmLuna.frx":220348
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   17
      Left            =   12240
      Picture         =   "frmLuna.frx":222012
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   16
      Left            =   11520
      Picture         =   "frmLuna.frx":223CDC
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   15
      Left            =   10800
      Picture         =   "frmLuna.frx":2259A6
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   14
      Left            =   10080
      Picture         =   "frmLuna.frx":227670
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   13
      Left            =   9360
      Picture         =   "frmLuna.frx":22933A
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   12
      Left            =   8640
      Picture         =   "frmLuna.frx":22B004
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   11
      Left            =   7920
      Picture         =   "frmLuna.frx":22CCCE
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   10
      Left            =   7200
      Picture         =   "frmLuna.frx":22E998
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   9
      Left            =   6480
      Picture         =   "frmLuna.frx":230662
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   8
      Left            =   5760
      Picture         =   "frmLuna.frx":23232C
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   7
      Left            =   5040
      Picture         =   "frmLuna.frx":233FF6
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   6
      Left            =   4320
      Picture         =   "frmLuna.frx":235CC0
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   5
      Left            =   3600
      Picture         =   "frmLuna.frx":23798A
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   4
      Left            =   2880
      Picture         =   "frmLuna.frx":239654
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   3
      Left            =   2160
      Picture         =   "frmLuna.frx":23B31E
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   2
      Left            =   1440
      Picture         =   "frmLuna.frx":23CFE8
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   1
      Left            =   720
      Picture         =   "frmLuna.frx":23ECB2
      Top             =   0
      Width           =   720
   End
   Begin VB.Image imgTile 
      Height          =   720
      Index           =   0
      Left            =   0
      Picture         =   "frmLuna.frx":24097C
      Top             =   0
      Width           =   720
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   1860
      Top             =   9420
      Width           =   855
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unit Commands"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1455
      Left            =   3120
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   9540
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   60
      Top             =   9420
      Width           =   1755
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TileState(0 To 259), BuildingState(0 To 259) As Integer
Dim TileOwner(0 To 259), BuildingHP(0 To 259), BuildingHealth(0 To 259)
Dim Firepower(), Range(), Movement(), Accuracy(), HP(), Mode(), UnitType(), UnitTeam(), HArmor(), PArmor(), CArmor(), Health(), UnitLoc(), AttackType()
Dim Units, ActiveUnit, MovesLeft As Single
Dim Turn
Dim BlueOre As Integer, BlueVespian As Integer, RedOre As Integer, RedVespian As Integer, _
    BlueBarracks As Integer, RedBarracks As Integer, BlueMechLab As Integer, _
    RedMechLab As Integer, BlueUnits As Integer, RedUnits As Integer, BlueSupport As Integer, _
    RedSupport As Integer
    Dim BlueHatchery As Integer, RedHatchery As Integer, BlueCryogenics As Integer, RedGeothermal As Integer

Private Sub cmdEscape_Click()
If MsgBox("Are you sure you want to exit Solaris 3000?", vbYesNo, "Escape Hatch Activated") = vbYes Then
Unload Me
End If
End Sub



Private Sub Form_Load()
Randomize
Dim i
For i = 0 To 259
If Int(Rnd * 3 + 1) = 1 Then
imgTile(i).Picture = imgPlains.Picture
TileState(i) = "Plains"
End If
    If i > 129 Then
    If Int(Rnd * 10 + 1) = 1 Then
    imgTile(i).Picture = imgLava.Picture
    TileState(i) = "Lava"
    End If
    Else
    If Int(Rnd * 36 + 1) = 1 Then
    imgTile(i).Picture = imgLava.Picture
    TileState(i) = "Lava"
    End If
    End If
    If i < 130 Then
    If Int(Rnd * 10 + 1) = 1 And TileState(i) = "" Then
    imgTile(i).Picture = imgGlacier.Picture
    TileState(i) = "Glacier"
    End If
    Else
    If Int(Rnd * 36 + 1) = 1 And TileState(i) = "" Then
    imgTile(i).Picture = imgGlacier.Picture
    TileState(i) = "Glacier"
    End If
    End If
If Int(Rnd * 4 + 1) = 1 Then
imgTile(i).Picture = imgOre.Picture
TileState(i) = "Ore"
End If
If Int(Rnd * 30 + 1) = 1 Then
imgTile(i).Picture = imgVespian.Picture
TileState(i) = "Vespian"
End If
Next i
imgTile(44).Picture = imgVespian.Picture
TileState(44) = "Vespian"
imgTile(215).Picture = imgVespian.Picture
TileState(215) = "Vespian"
imgTile(0).Picture = imgBlueTownHall.Picture
TileState(0) = "Town Hall"
TileOwner(0) = "Blue"
BuildingHP(0) = 5000
BuildingHealth(0) = 5000
imgTile(259).Picture = imgRedTownHall.Picture
TileState(259) = "Town Hall"
TileOwner(259) = "Red"
BuildingHP(259) = 5000
BuildingHealth(259) = 5000
Call LoadUnit("Worker", "Blue", 0, 1, 0, 1, 1, 100, 0.1, 0.1, 0.1, "None")
Call LoadUnit("Worker", "Red", 259, 1, 0, 1, 1, 100, 0.1, 0.1, 0.1, "None")
Turn = "Blue"
ActiveUnit = 1
BlueOre = 10
RedOre = 9
BlueVespian = 5
RedVespian = 4
MovesLeft = 1
BlueUnits = 1
BlueSupport = 3
RedUnits = 1
RedSupport = 3
Me.MouseIcon = imgUnitMode(0).Picture
Me.BackColor = RGB(10, 200, 200)
lblOre.Caption = BlueOre
lblVespian.Caption = BlueVespian
lblSupport.Caption = BlueUnits & "/" & BlueSupport
'Call TurnCycle
End Sub


Private Sub LoadUnit(Unit As String, Team As String, Location As Integer, UnitAccuracy, _
UnitFirepower, UnitRange, UnitMovement, UnitHP, UnitHackArmor, UnitPierceArmor, UnitCrushArmor, UnitAttack)
Units = Units + 1
ReDim Preserve Firepower(1 To Units), Range(1 To Units), Movement(1 To Units), _
Accuracy(1 To Units), HP(1 To Units), Mode(1 To Units), UnitType(1 To Units), _
UnitTeam(1 To Units), Health(1 To Units), UnitLoc(1 To Units), HArmor(1 To Units), PArmor(1 To Units), _
CArmor(1 To Units), AttackType(1 To Units)
Load imgUnit(Units)
With imgUnit(Units)
.Visible = True
.Top = imgTile(Location).Top
.Left = imgTile(Location).Left
.ZOrder
End With
Firepower(Units) = UnitFirepower
Range(Units) = UnitRange
Movement(Units) = UnitMovement
Accuracy(Units) = UnitAccuracy
HP(Units) = UnitHP
Health(Units) = HP(Units)
UnitLoc(Units) = Location
Mode(Units) = "Move"
UnitType(Units) = Unit
UnitTeam(Units) = Team
HArmor(Units) = UnitHackArmor
PArmor(Units) = UnitPierceArmor
CArmor(Units) = UnitCrushArmor
AttackType(Units) = UnitAttack
Select Case Unit
Case Is = "Worker"
If UnitTeam(Units) = "Red" Then
imgUnit(Units).Picture = imgWorkerRed.Picture
Else
imgUnit(Units).Picture = imgWorkerBlue.Picture
End If
Case Is = "Trooper"
If UnitTeam(Units) = "Red" Then
imgUnit(Units).Picture = imgTrooperRed.Picture
Else
imgUnit(Units).Picture = imgTrooperBlue.Picture
End If
Case Is = "Jedi"
If UnitTeam(Units) = "Red" Then
imgUnit(Units).Picture = imgJediRed.Picture
Else
imgUnit(Units).Picture = imgJediBlue.Picture
End If
Case Is = "Tank"
If UnitTeam(Units) = "Red" Then
imgUnit(Units).Picture = imgTankRed.Picture
Else
imgUnit(Units).Picture = imgTankBlue.Picture
End If
Case Is = "Gribble"
If UnitTeam(Units) = "Red" Then
imgUnit(Units).Picture = imgGribbleRed.Picture
Else
imgUnit(Units).Picture = imgGribbleBlue.Picture
End If
Case Is = "Cryomegalon"
imgUnit(Units).Picture = imgCryomegalon.Picture
Case Is = "Pyromegalon"
imgUnit(Units).Picture = imgPyromegalon.Picture
End Select

End Sub
Private Sub TurnCycle()
Dim i
'''Check Victory Conditions...
If BlueUnits = 0 Then MsgBox "Red Wins!", , "Victory"
If RedUnits = 0 Then MsgBox "Blue Wins!", , "Victory"
'For i = 1 To Units
'If UnitTeam(i) = "Red" And Mode(i) <> "Dead" Then GoTo 1:
'Next i
'MsgBox "Blue Wins!", , "Victory"
'1:
'For i = 1 To Units
'If UnitTeam(i) = "Blue" And Mode(i) <> "Dead" Then GoTo 2:
'Next i
'MsgBox "Red Wins!", , "Victory"
'2:
'''Cycle Thru Units
Repeat:
ActiveUnit = ActiveUnit + 1
If ActiveUnit > Units Then
ActiveUnit = 1
Call SwitchPlayer
End If
If UnitTeam(ActiveUnit) = Turn And Mode(ActiveUnit) <> "Dead" Then
GoTo 3
Else
GoTo Repeat
End If
3:
imgActiveUnit.Picture = imgUnit(ActiveUnit).Picture
linHealthBar.X2 = linHealthBar.X1 + 720 * (Health(ActiveUnit) / HP(ActiveUnit))
linHealthBar.BorderColor = RGB(255 - (255 * (Health(ActiveUnit) / HP(ActiveUnit))), 255 * (Health(ActiveUnit) / HP(ActiveUnit)), 0)
Select Case UnitType(ActiveUnit)
Case Is = "Worker"
imgUnitMode(1).Visible = True
imgUnitMode(3).Visible = False
Case Is = "Trooper", "Jedi", "Tank", "Gribble", "Pyromegalon", "Cryomegalon"
imgUnitMode(1).Visible = False
imgUnitMode(3).Visible = True
End Select
Mode(ActiveUnit) = "Move"
frmMain.MouseIcon = imgUnitMode(0).Picture
imgHighlight.Move imgUnit(ActiveUnit).Left, imgUnit(ActiveUnit).Top
MovesLeft = Movement(ActiveUnit)
lblActiveUnit.Caption = UnitTeam(ActiveUnit) & " " & UnitType(ActiveUnit) & " HP " & Health(ActiveUnit) & "/" & HP(ActiveUnit) & " Moves: " & MovesLeft
imgUnit(ActiveUnit).ZOrder
End Sub

Private Sub Image2_Click()

End Sub

Private Sub imgBuilding_Click(Index As Integer)
Select Case Index
Case Is = 0
Mode(ActiveUnit) = "Ore Mine"
Case Is = 1
Mode(ActiveUnit) = "Vespian Mine"
Case Is = 2
Mode(ActiveUnit) = "Barracks"
Case Is = 3
Mode(ActiveUnit) = "Defense Tower"
Case Is = 4
Mode(ActiveUnit) = "Mech Lab"
Case Is = 5
Mode(ActiveUnit) = "Life Support"
Case Is = 6
Mode(ActiveUnit) = "Geothermal Plant"
Case Is = 7
Mode(ActiveUnit) = "Gribble Hatchery"
Case Is = 8
Mode(ActiveUnit) = "Cryogenics"
End Select
frmMain.MouseIcon = imgUnitMode(1).Picture
End Sub

Private Sub imgBuildUnit_Click(Index As Integer)
Select Case Index
Case Is = 0 ' Worker
If Turn = "Blue" Then
Call LoadUnit("Worker", "Blue", 0, 1, 0, 1, 1, 100, 0.1, 0.1, 0.1, "None")
BlueOre = BlueOre - 10
lblOre.Caption = BlueOre
Else
Call LoadUnit("Worker", "Red", 259, 1, 0, 1, 1, 85, 0.1, 0.1, 0.25, "None")
RedOre = RedOre - 10
lblOre.Caption = RedOre
End If
Case Is = 1 ' Trooper
If Turn = "Blue" Then
Call LoadUnit("Trooper", "Blue", 0, 4 / 5, 50, 2, 2, 150, 0.2, 0.4, 0.2, "Pierce")
BlueOre = BlueOre - 20
lblOre.Caption = BlueOre
Else
Call LoadUnit("Trooper", "Red", 259, 4 / 5, 65, 1, 2, 135, 0.2, 0.4, 0.2, "Pierce")
RedOre = RedOre - 20
lblOre.Caption = RedOre
End If
Case Is = 2 ' Jedi
If Turn = "Blue" Then
Call LoadUnit("Jedi", "Blue", 0, 7 / 8, 80, 1, 2, 200, 0.3, 0.3, 0.2, "Hack")
BlueOre = BlueOre - 30
lblOre.Caption = BlueOre
Else
Call LoadUnit("Jedi", "Red", 259, 7 / 8, 90, 1, 2, 180, 0.3, 0.3, 0.2, "Hack")
RedOre = RedOre - 30
lblOre.Caption = RedOre
End If
Case Is = 3 ' tank
If Turn = "Blue" Then
Call LoadUnit("Tank", "Blue", 0, 4 / 5, 100, 3, 2, 200, 0.9, 0, 0.9, "Crush")
BlueOre = BlueOre - 40
lblOre.Caption = BlueOre
Else
Call LoadUnit("Tank", "Red", 259, 4 / 5, 115, 3, 2, 210, 0.9, 0, 0.75, "Crush")
RedOre = RedOre - 40
lblOre.Caption = RedOre
End If
Case Is = 4 ' Gribble
If Turn = "Blue" Then
Call LoadUnit("Gribble", "Blue", 0, 9 / 10, 160, 1, 3, 200, 0.65, 0.9, 0.9, "Crush")
BlueOre = BlueOre - 80
lblOre.Caption = BlueOre
Else
Call LoadUnit("Gribble", "Red", 259, 9 / 10, 180, 1, 3, 220, 0.5, 0.9, 0.8, "Crush")
RedOre = RedOre - 80
lblOre.Caption = RedOre
End If
Case Is = 5 ' Pyromegalon
Call LoadUnit("Pyromegalon", "Red", 259, 7 / 8, 90, 2, 3, 300, 0.2, 0.2, 0.6, "Burn")
RedVespian = RedVespian - 18
lblVespian.Caption = RedVespian
Case Is = 6 ' Cryomegalon
Call LoadUnit("Cryomegalon", "Blue", 0, 9 / 10, 115, 1, 3, 300, 0.4, 0.25, 0.4, "Pierce")
BlueOre = BlueOre - 50
lblOre.Caption = BlueOre
End Select
If Turn = "Blue" Then
BlueUnits = BlueUnits + 1
lblSupport.Caption = BlueUnits & "/" & BlueSupport
Else
RedUnits = RedUnits + 1
lblSupport.Caption = RedUnits & "/" & RedSupport
End If
Call CheckConstruction
End Sub

Private Sub imgTile_Click(Index As Integer)
Dim i
Select Case Mode(ActiveUnit)
Case Is = "Move"
If TileState(Index) <> "Geothermal Plant" And TileState(Index) <> "Lava" Then
If TileState(Index) = "Town Hall" And TileState(Index) <> Turn Then Exit Sub
If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 And _
imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 Then
If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 And _
imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 Then
imgUnit(ActiveUnit).Move imgUnit(ActiveUnit).Left, imgTile(Index).Top
imgUnit(ActiveUnit).Move imgTile(Index).Left, imgUnit(ActiveUnit).Top
imgHighlight.Move imgUnit(ActiveUnit).Left, imgUnit(ActiveUnit).Top
If TileState(Index) = "Plains" Then
MovesLeft = MovesLeft - 0.5
Else
MovesLeft = MovesLeft - 1
End If
lblActiveUnit.Caption = UnitTeam(ActiveUnit) & " " & UnitType(ActiveUnit) & " HP " & Health(ActiveUnit) & "/" & HP(ActiveUnit) & " Moves: " & MovesLeft
UnitLoc(ActiveUnit) = Index
On Error Resume Next
Call TowerShoot(ActiveUnit)
If MovesLeft <= 0 And Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
End If
End If
End If
Case Is = "Ore Mine"
    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If imgTile(Index).Picture = imgOre.Picture Then
            If Turn = "Blue" Then
            imgTile(Index).Picture = imgConstruct1Blue.Picture
            BlueVespian = BlueVespian - 2
            BlueOre = BlueOre - 2
            lblVespian.Caption = BlueVespian
            lblOre.Caption = BlueOre
            Else '''Red's Turn
            imgTile(Index).Picture = imgConstruct1Red.Picture
            RedVespian = RedVespian - 2
            RedOre = RedOre - 2
            lblVespian.Caption = RedVespian
            lblOre.Caption = RedOre
            End If
            TileOwner(Index) = Turn
            TileState(Index) = "Ore Mine"
            BuildingState(Index) = -1
            Call TowerShoot(ActiveUnit)
            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
            End If
        End If
    End If
Case Is = "Geothermal Plant"
    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If imgTile(Index).Picture = imgLava.Picture Then
            If Turn = "Blue" Then
            imgTile(Index).Picture = imgConstruct1Blue.Picture
            BlueVespian = BlueVespian - 20
            BlueOre = BlueOre - 5
            lblVespian.Caption = BlueVespian
            lblOre.Caption = BlueOre
            Else '''Red's Turn
            imgTile(Index).Picture = imgConstruct1Red.Picture
            RedVespian = RedVespian - 20
            RedOre = RedOre - 5
            lblVespian.Caption = RedVespian
            lblOre.Caption = RedOre
            End If
            TileOwner(Index) = Turn
            TileState(Index) = "Geothermal Plant"
            BuildingState(Index) = -5
            On Error Resume Next
            Call TowerShoot(ActiveUnit)
            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
            End If
        End If
    End If
Case Is = "Cryogenics"
    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If imgTile(Index).Picture = imgGlacier.Picture Then
            imgTile(Index).Picture = imgConstruct1Blue.Picture
            BlueVespian = BlueVespian - 20
            BlueOre = BlueOre - 5
            lblVespian.Caption = BlueVespian
            lblOre.Caption = BlueOre
            TileOwner(Index) = Turn
            TileState(Index) = "Cryogenics Lab"
            BuildingState(Index) = -5
            On Error Resume Next
            Call TowerShoot(ActiveUnit)
            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
            End If
        End If
    End If
Case Is = "Life Support"
    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If Turn = "Blue" Then
            imgTile(Index).Picture = imgConstruct1Blue.Picture
            BlueVespian = BlueVespian - 5
            BlueOre = BlueOre - 2
            lblSupport.Caption = BlueUnits & "/" & BlueSupport
            lblVespian.Caption = BlueVespian
            lblOre.Caption = BlueOre
            Else '''Red's Turn
            imgTile(Index).Picture = imgConstruct1Red.Picture
            RedVespian = RedVespian - 5
            RedOre = RedOre - 2
            lblSupport.Caption = RedUnits & "/" & RedSupport
            lblVespian.Caption = RedVespian
            lblOre.Caption = RedOre
            End If
            TileOwner(Index) = Turn
            TileState(Index) = "Life Support System"
            BuildingState(Index) = -2
            On Error Resume Next
            Call TowerShoot(ActiveUnit)
            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
        End If
    End If
Case Is = "Vespian Mine"
    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If imgTile(Index).Picture = imgVespian.Picture Then
            If Turn = "Blue" Then
            imgTile(Index).Picture = imgConstruct1Blue.Picture
            BlueVespian = BlueVespian - 2
            BlueOre = BlueOre - 2
            lblVespian.Caption = BlueVespian
            lblOre.Caption = BlueOre
            Else '''Red's Turn
            imgTile(Index).Picture = imgConstruct1Red.Picture
            RedVespian = RedVespian - 2
            RedOre = RedOre - 2
            lblVespian.Caption = RedVespian
            lblOre.Caption = RedOre
            End If
            TileOwner(Index) = Turn
            TileState(Index) = "Vespian Mine"
            BuildingState(Index) = -3
            On Error Resume Next
            Call TowerShoot(ActiveUnit)
            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
            End If
        End If
    End If
Case Is = "Barracks"
    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If Turn = "Blue" Then
            imgTile(Index).Picture = imgConstruct1Blue.Picture
            BlueVespian = BlueVespian - 10
            BlueOre = BlueOre - 5
            lblVespian.Caption = BlueVespian
            lblOre.Caption = BlueOre
            Else '''Red's Turn
            imgTile(Index).Picture = imgConstruct1Red.Picture
            RedVespian = RedVespian - 10
            RedOre = RedOre - 5
            lblVespian.Caption = RedVespian
            lblOre.Caption = RedOre
            End If
            TileOwner(Index) = Turn
            TileState(Index) = "Barracks"
            BuildingState(Index) = -7
            On Error Resume Next
            Call TowerShoot(ActiveUnit)
            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
        End If
    End If
    Case Is = "Mech Lab"
    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If Turn = "Blue" Then
            imgTile(Index).Picture = imgConstruct1Blue.Picture
            BlueVespian = BlueVespian - 20
            BlueOre = BlueOre - 5
            lblVespian.Caption = BlueVespian
            lblOre.Caption = BlueOre
            Else '''Red's Turn
            imgTile(Index).Picture = imgConstruct1Red.Picture
            RedVespian = RedVespian - 20
            RedOre = RedOre - 5
            lblVespian.Caption = RedVespian
            lblOre.Caption = RedOre
            End If
            TileOwner(Index) = Turn
            TileState(Index) = "Mech Lab"
            BuildingState(Index) = -7
            On Error Resume Next
            Call TowerShoot(ActiveUnit)
            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
        End If
    End If

    Case Is = "Gribble Hatchery"
    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If Turn = "Blue" Then
            imgTile(Index).Picture = imgConstruct1Blue.Picture
            BlueVespian = BlueVespian - 40
            BlueOre = BlueOre - 10
            lblVespian.Caption = BlueVespian
            lblOre.Caption = BlueOre
            Else '''Red's Turn
            imgTile(Index).Picture = imgConstruct1Red.Picture
            RedVespian = RedVespian - 40
            RedOre = RedOre - 10
            lblVespian.Caption = RedVespian
            lblOre.Caption = RedOre
            End If
            TileOwner(Index) = Turn
            TileState(Index) = "Gribble Hatchery"
            BuildingState(Index) = -9
            On Error Resume Next
            Call TowerShoot(ActiveUnit)
            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
        End If
    End If
Case Is = "Defense Tower"
    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If Turn = "Blue" Then
            imgTile(Index).Picture = imgConstruct1Blue.Picture
            BlueVespian = BlueVespian - 10
            BlueOre = BlueOre - 5
            lblVespian.Caption = BlueVespian
            lblOre.Caption = BlueOre
            Else '''Red's Turn
            imgTile(Index).Picture = imgConstruct1Red.Picture
            RedVespian = RedVespian - 10
            RedOre = RedOre - 5
            lblVespian.Caption = RedVespian
            lblOre.Caption = RedOre
            End If
            TileOwner(Index) = Turn
            TileState(Index) = "Defense Tower"
            BuildingState(Index) = -7
            On Error Resume Next
            Call TowerShoot(ActiveUnit)
            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
        End If
    End If
Case Is = "Build"
    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If TileOwner(Index) = Turn And BuildingState(Index) < 0 Then
            BuildingState(Index) = BuildingState(Index) + 1
                If BuildingState(Index) = 0 Then Call BuildingComplete(Index)
                On Error Resume Next
                Call TowerShoot(ActiveUnit)
            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
            End If
        End If
    End If
Case Is = "Attack"
Call BuildingAttack(Index)
'    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
'    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
'        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
'        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
'            If TileOwner(Index) <> Turn And BuildingState(Index) = 0 Then
'            If Rnd <= Accuracy(ActiveUnit) Then BuildingHealth(Index) = BuildingHealth(Index) - Firepower(ActiveUnit)
'                If BuildingHealth(Index) <= 0 Then
'                BuildingHealth(Index) = 0
'                    If TileState(Index) = "Barracks" Then
'                        If Turn = "Blue" Then
'                        'Kill Red's Barracks + Supported Building
'                        RedBarracks = RedBarracks - 1
'                        If RedBarracks = 0 Then
'                        For i = 0 To 259
'                            If TileState(i) = "Mech Lab" And TileOwner(i) <> Turn Then
'                            imgTile(i).Picture = imgLuna.Picture
'                            TileOwner(i) = ""
'                            TileState(i) = ""
'                            End If
'                            If TileState(i) = "Gribble Hatchery" And TileOwner(i) <> Turn Then
'                            imgTile(i).Picture = imgLuna.Picture
'                            TileOwner(i) = ""
'                            TileState(i) = ""
'                            End If
'                            Next i
'                            RedMechLab = 0
'                            RedHatchery = 0
'                            End If
'                        Else
'                        'Kill Blue's Barracks etc.
'                        BlueBarracks = BlueBarracks - 1
'                        If BlueBarracks = 0 Then
'                        For i = 0 To 259
'                            If TileState(i) = "Mech Lab" And TileOwner(i) <> Turn Then
'                            imgTile(i).Picture = imgLuna.Picture
'                            TileOwner(i) = ""
'                            TileState(i) = ""
'                            End If
'                            If TileState(i) = "Gribble Hatchery" And TileOwner(i) <> Turn Then
'                            imgTile(i).Picture = imgLuna.Picture
'                            TileOwner(i) = ""
'                            TileState(i) = ""
'                            End If
'                        Next i
'                        BlueMechLab = 0
'                        BlueHatchery = 0
'                        End If
'                        End If
'                    End If
'                If TileState(Index) = "Mech Lab" And BuildingHealth(Index) <= 0 Then
'                If Turn = "Blue" Then
'                BlueMechLab = BlueMechLab - 1
'                Else
'                RedMechLab = RedMechLab - 1
'                End If
'                End If
'                TileState(Index) = ""
'                TileOwner(Index) = ""
'                imgTile(Index).Picture = imgLuna.Picture
'                End If
'                If TileState(Index) = "Gribble Hatchery" And BuildingHealth(Index) <= 0 Then
'                If Turn = "Blue" Then
'                BlueHatchery = BlueHatchery - 1
'                Else
'                RedHatchery = RedHatchery - 1
'                End If
'                TileState(Index) = ""
'                TileOwner(Index) = ""
'                imgTile(Index).Picture = imgLuna.Picture
'                End If
'                If TileState(Index) = "Geothermal Plant" And BuildingHealth(Index) <= 0 Then
'                RedGeothermal = RedGeothermal - 1
'                End If
'                TileState(Index) = ""
'                TileOwner(Index) = ""
'                imgTile(Index).Picture = imgLuna.Picture
'                If TileState(Index) = "Cryogenics Lab" And BuildingHealth(Index) <= 0 Then
'                BlueCryogenics = BlueCryogenics - 1
'                End If
'                TileState(Index) = ""
'                TileOwner(Index) = ""
'                imgTile(Index).Picture = imgLuna.Picture
'                On Error Resume Next
'                Call TowerShoot(ActiveUnit)
'            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
'            End If
'            End If
''            If TileOwner(Index) <> Turn And BuildingState(Index) < 0 Then
''            TileState(Index) = ""
''            TileOwner(Index) = ""
''            imgTile(Index).Picture = imgLuna.Picture
'            On Error Resume Next
'            Call TowerShoot(ActiveUnit)
'            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
'            End If
'        End If
    End Select
    CheckConstruction
End Sub


Private Sub imgTile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgViewBuilding.Picture = imgTile(Index).Picture
If BuildingHP(Index) = "" Or BuildingHealth(Index) <= 0 Then
linBldgHealthBar.Visible = False
Else
linBldgHealthBar.Visible = True
linBldgHealthBar.X2 = linBldgHealthBar.X1 + 720 * (BuildingHealth(Index) / BuildingHP(Index))
linBldgHealthBar.BorderColor = RGB(255 - (255 * (BuildingHealth(Index) / BuildingHP(Index))), 255 * (BuildingHealth(Index) / BuildingHP(Index)), 0)
End If
linViewUnitHealthBar.Visible = False
imgViewUnit.Visible = False
lblViewUnit.Caption = ""
lblViewBuilding.Caption = TileOwner(Index) & " " & TileState(Index)
If BuildingState(Index) = 0 And TileOwner(Index) <> "" Then lblViewBuilding.Caption = _
    lblViewBuilding.Caption & " " & BuildingHealth(Index) & "/" & BuildingHP(Index) & " HP"
If BuildingState(Index) < 0 And TileOwner(Index) <> "" Then lblViewBuilding.Caption = _
    lblViewBuilding.Caption & " " & -BuildingState(Index) & " To Build"
If TileState(Index) = "" Then lblViewBuilding.Caption = "Rocky Outcropping"
End Sub

Private Sub imgUnit_Click(Index As Integer)
If Mode(ActiveUnit) = "Attack" And UnitTeam(Index) <> Turn Then
    If imgUnit(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgUnit(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgUnit(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgUnit(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If Rnd <= Accuracy(ActiveUnit) Then
                Select Case AttackType(ActiveUnit)
                Case Is = "Hack"
                Health(Index) = Health(Index) - Firepower(ActiveUnit) * (1 - HArmor(Index))
                Case Is = "Pierce"
                Health(Index) = Health(Index) - Firepower(ActiveUnit) * (1 - PArmor(Index))
                Case Is = "Crush"
                Health(Index) = Health(Index) - Firepower(ActiveUnit) * (1 - CArmor(Index))
                Case Else
                Health(Index) = Health(Index) - Firepower(ActiveUnit)
                End Select
                If Health(Index) <= 0 Then
                Mode(Index) = "Dead"
                    If Turn = "Blue" Then
                    RedUnits = RedUnits - 1
                    Else
                    BlueUnits = BlueUnits - 1
                    End If
                    If Health(Index) < 0 Then Health(Index) = 0
                    imgUnit(Index).Top = -720
                    End If
                End If
        Call TurnCycle
        End If
    End If
End If
CheckConstruction
End Sub

Private Sub imgUnit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgViewBuilding.Picture = imgTile(UnitLoc(Index)).Picture
imgViewUnit.Visible = True
imgViewUnit.Picture = imgUnit(Index).Picture
If BuildingHP(UnitLoc(Index)) = "" Or BuildingHealth(UnitLoc(Index)) <= 0 Then
linBldgHealthBar.Visible = False
Else
linBldgHealthBar.Visible = True
linBldgHealthBar.X2 = linBldgHealthBar.X1 + 720 * (BuildingHealth(UnitLoc(Index)) / BuildingHP(UnitLoc(Index)))
linBldgHealthBar.BorderColor = RGB(255 - (255 * (BuildingHealth(UnitLoc(Index)) / BuildingHP(UnitLoc(Index)))), 255 * (BuildingHealth(UnitLoc(Index)) / BuildingHP(UnitLoc(Index))), 0)
End If
lblViewBuilding.Caption = TileOwner(UnitLoc(Index)) & " " & TileState(UnitLoc(Index))
If BuildingState(UnitLoc(Index)) = 0 And TileOwner(UnitLoc(Index)) <> "" Then lblViewBuilding.Caption = _
    lblViewBuilding.Caption & " " & BuildingHealth(UnitLoc(Index)) & "/" & BuildingHP(UnitLoc(Index)) & " HP"
If BuildingState(UnitLoc(Index)) < 0 And TileOwner(UnitLoc(Index)) <> "" Then lblViewBuilding.Caption = _
    lblViewBuilding.Caption & " " & -BuildingState(UnitLoc(Index)) & " To Build"
If TileState(UnitLoc(Index)) = "" Then lblViewBuilding.Caption = "Rocky Outcropping"
If Health(Index) > 0 Then
linViewUnitHealthBar.Visible = True
linViewUnitHealthBar.X2 = linViewUnitHealthBar.X1 + 720 * (Health(Index) / HP(Index))
linViewUnitHealthBar.BorderColor = RGB(255 - (255 * (Health(Index) / HP(Index))), 255 * (Health(Index) / HP(Index)), 0)
lblViewUnit.Caption = UnitTeam(Index) & " " & UnitType(Index) & " " & Health(Index) & "/" & HP(Index) & " HP"
End If
End Sub

Private Sub imgUnitMode_Click(Index As Integer)
Select Case Index
Case Is = 0
Mode(ActiveUnit) = "Move"
Case Is = 1
Mode(ActiveUnit) = "Build"
Case Is = 2
If Health(ActiveUnit) < HP(ActiveUnit) Then
Health(ActiveUnit) = Health(ActiveUnit) + (Int(HP(ActiveUnit)) / 5)
If Health(ActiveUnit) > HP(ActiveUnit) Then Health(ActiveUnit) = HP(ActiveUnit)
End If
Call TurnCycle
Call CheckConstruction
Case Is = 3
Mode(ActiveUnit) = "Attack"
End Select
If Index = 0 Or Index = 1 Or Index = 3 Then frmMain.MouseIcon = imgUnitMode(Index).Picture
End Sub

Private Sub SwitchPlayer()
Dim OreGain, VespianGain, i
    If Turn = "Red" Then
    Turn = "Blue"
    imgBuilding(0).Picture = imgOreMineBlue.Picture
    imgBuilding(1).Picture = imgVespianMineBlue.Picture
    imgBuilding(2).Picture = imgBarracksBlue.Picture
    imgBuilding(3).Picture = imgDefenseTowerBlue.Picture
    imgBuilding(4).Picture = imgMechLabBlue.Picture
    imgBuilding(5).Picture = imgLifeSupportBlue.Picture
    imgBuilding(7).Picture = imgHatcheryBlue.Picture
    Me.BackColor = RGB(10, 200, 200)
    imgBuildUnit(0).Picture = imgWorkerBlue.Picture
    imgBuildUnit(1).Picture = imgTrooperBlue.Picture
    imgBuildUnit(2).Picture = imgJediBlue.Picture
    imgBuildUnit(3).Picture = imgTankBlue.Picture
    imgBuildUnit(4).Picture = imgGribbleBlue.Picture
    Else
    Turn = "Red"
    imgBuilding(0).Picture = imgOreMineRed.Picture
    imgBuilding(1).Picture = imgVespianMineRed.Picture
    imgBuilding(2).Picture = imgBarracksRed.Picture
    imgBuilding(3).Picture = imgDefenseTowerRed.Picture
    imgBuilding(4).Picture = imgMechLabRed.Picture
    imgBuilding(5).Picture = imgLifeSupportRed.Picture
    imgBuilding(7).Picture = imgHatcheryRed.Picture
    Me.BackColor = RGB(200, 10, 10)
    imgBuildUnit(0).Picture = imgWorkerRed.Picture
    imgBuildUnit(1).Picture = imgTrooperRed.Picture
    imgBuildUnit(2).Picture = imgJediRed.Picture
    imgBuildUnit(3).Picture = imgTankRed.Picture
    imgBuildUnit(4).Picture = imgGribbleRed.Picture
    End If
    OreGain = 1
    VespianGain = 1
    For i = 0 To 259
    If TileOwner(i) = Turn And TileState(i) = "Ore Mine" And BuildingState(i) = 0 Then
    OreGain = OreGain + 1
    End If
    If TileOwner(i) = Turn And TileState(i) = "Vespian Mine" And BuildingState(i) = 0 Then
    VespianGain = VespianGain + 1
    End If
    If TileOwner(i) = Turn And TileState(i) = "Geothermal Plant" And BuildingState(i) = 0 Then
    VespianGain = VespianGain + 3
    End If
    If TileOwner(i) = Turn And TileState(i) = "Cryogenics Lab" And BuildingState(i) = 0 Then
    OreGain = OreGain + 8
    End If
    Next i
    Select Case Turn
    Case Is = "Red"
    RedOre = RedOre + OreGain
    lblOre.Caption = RedOre
    RedVespian = RedVespian + VespianGain
    lblVespian.Caption = RedVespian
    lblSupport.Caption = RedUnits & "/" & RedSupport
    Case Is = "Blue"
    BlueOre = BlueOre + OreGain
    lblOre.Caption = BlueOre
    BlueVespian = BlueVespian + VespianGain
    lblVespian.Caption = BlueVespian
    lblSupport.Caption = BlueUnits & "/" & BlueSupport
    End Select
    CheckConstruction
End Sub

Private Sub BuildingComplete(Tile)
            Select Case TileState(Tile)
            Case Is = "Life Support System"
            If TileOwner(Tile) = "Blue" Then
            imgTile(Tile).Picture = imgLifeSupportBlue.Picture
            BlueSupport = BlueSupport + 3
            lblSupport.Caption = BlueUnits & "/" & BlueSupport
            Else
            imgTile(Tile).Picture = imgLifeSupportRed.Picture
            RedSupport = RedSupport + 3
            lblSupport.Caption = RedUnits & "/" & RedSupport
            End If
            BuildingHP(Tile) = 550
            BuildingHealth(Tile) = 550
            Case Is = "Ore Mine"
            If TileOwner(Tile) = "Blue" Then
            imgTile(Tile).Picture = imgOreMineBlue.Picture
            Else
            imgTile(Tile).Picture = imgOreMineRed.Picture
            End If
            BuildingHP(Tile) = 500
            BuildingHealth(Tile) = 500
            Case Is = "Vespian Mine"
            If TileOwner(Tile) = "Blue" Then
            imgTile(Tile).Picture = imgVespianMineBlue.Picture
            Else
            imgTile(Tile).Picture = imgVespianMineRed.Picture
            End If
            BuildingHP(Tile) = 500
            BuildingHealth(Tile) = 500
            Case Is = "Geothermal Plant"
            imgTile(Tile).Picture = imgGeothermalRed.Picture
            BuildingHP(Tile) = 860
            BuildingHealth(Tile) = 860
            RedGeothermal = RedGeothermal + 1
            Case Is = "Cryogenics Lab"
            imgTile(Tile).Picture = imgCryogenicsBlue.Picture
            BuildingHP(Tile) = 880
            BuildingHealth(Tile) = 880
            BlueCryogenics = BlueCryogenics + 1
            Case Is = "Barracks"
            If TileOwner(Tile) = "Blue" Then
            imgTile(Tile).Picture = imgBarracksBlue.Picture
            BlueBarracks = BlueBarracks + 1
            Else
            imgTile(Tile).Picture = imgBarracksRed.Picture
            RedBarracks = RedBarracks + 1
            End If
            BuildingHP(Tile) = 1500
            BuildingHealth(Tile) = 1500
            Case Is = "Defense Tower"
            If TileOwner(Tile) = "Blue" Then
            imgTile(Tile).Picture = imgDefenseTowerBlue.Picture
            Else
            imgTile(Tile).Picture = imgDefenseTowerRed.Picture
            End If
            BuildingHP(Tile) = 2500
            BuildingHealth(Tile) = 2500
            Case Is = "Mech Lab"
            If TileOwner(Tile) = "Blue" Then
            imgTile(Tile).Picture = imgMechLabBlue.Picture
            BlueMechLab = BlueMechLab + 1
            Else
            imgTile(Tile).Picture = imgMechLabRed.Picture
            RedMechLab = RedMechLab + 1
            End If
            BuildingHP(Tile) = 1500
            BuildingHealth(Tile) = 1500
            Case Is = "Gribble Hatchery"
            If TileOwner(Tile) = "Blue" Then
            imgTile(Tile).Picture = imgHatcheryBlue.Picture
            BlueHatchery = BlueHatchery + 1
            Else
            imgTile(Tile).Picture = imgHatcheryRed.Picture
            RedHatchery = RedHatchery + 1
            End If
            BuildingHP(Tile) = 2000
            BuildingHealth(Tile) = 2000
            End Select
    If Turn = "Blue" Then
    lblSupport.Caption = BlueUnits & "/" & BlueSupport
    Else
    lblSupport.Caption = RedUnits & "/" & RedSupport
    End If
End Sub

Private Sub CheckConstruction()
If Turn = "Blue" Then
    If BlueOre < 2 Or BlueVespian < 2 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(0).Visible = False
    imgBuilding(1).Visible = False
    Else
    imgBuilding(0).Visible = True
    imgBuilding(1).Visible = True
    End If
    If BlueOre < 2 Or BlueVespian < 5 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(5).Visible = False
    Else
    imgBuilding(5).Visible = True
    End If
    If BlueOre < 5 Or BlueVespian < 10 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(2).Visible = False
    imgBuilding(3).Visible = False
    Else
    imgBuilding(2).Visible = True
    imgBuilding(3).Visible = True
    End If
    If BlueOre < 5 Or BlueVespian < 20 Or BlueBarracks = 0 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(4).Visible = False
    Else
    imgBuilding(4).Visible = True
    End If
    If BlueOre < 10 Or BlueVespian < 40 Or BlueBarracks = 0 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(7).Visible = False
    Else
    imgBuilding(7).Visible = True
    End If
    If BlueOre < 5 Or BlueVespian < 20 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(8).Visible = False
    Else
    imgBuilding(8).Visible = True
    End If
    imgBuilding(6).Visible = False
    '''Worker
    If BlueOre < 10 Or BlueUnits = BlueSupport Then
    imgBuildUnit(0).Visible = False
    Else
    imgBuildUnit(0).Visible = True
    End If
    '''Trooper
    If BlueOre < 20 Or BlueBarracks = 0 Or BlueUnits = BlueSupport Then
    imgBuildUnit(1).Visible = False
    Else
    imgBuildUnit(1).Visible = True
    End If
    '''Jedi
    If BlueOre < 30 Or BlueBarracks = 0 Or BlueUnits = BlueSupport Then
    imgBuildUnit(2).Visible = False
    Else
    imgBuildUnit(2).Visible = True
    End If
    '''Tank
    If BlueOre < 40 Or BlueMechLab = 0 Or BlueUnits = BlueSupport Then
    imgBuildUnit(3).Visible = False
    Else
    imgBuildUnit(3).Visible = True
    End If
    '''Gribble
    If BlueOre < 80 Or BlueHatchery = 0 Or BlueUnits = BlueSupport Then
    imgBuildUnit(4).Visible = False
    Else
    imgBuildUnit(4).Visible = True
    End If
    '''Cryomegalon
    If BlueOre < 50 Or BlueCryogenics = 0 Or BlueUnits = BlueSupport Then
    imgBuildUnit(6).Visible = False
    Else
    imgBuildUnit(6).Visible = True
    End If
    imgBuildUnit(5).Visible = False
Else
    If RedOre < 2 Or RedVespian < 2 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(0).Visible = False
    imgBuilding(1).Visible = False
    Else
    imgBuilding(0).Visible = True
    imgBuilding(1).Visible = True
    End If
    If RedOre < 2 Or RedVespian < 5 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(5).Visible = False
    Else
    imgBuilding(5).Visible = True
    End If
    If RedOre < 5 Or RedVespian < 10 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(2).Visible = False
    imgBuilding(3).Visible = False
    Else
    imgBuilding(2).Visible = True
    imgBuilding(3).Visible = True
    End If
    If RedOre < 5 Or RedVespian < 20 Or RedBarracks = 0 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(4).Visible = False
    Else
    imgBuilding(4).Visible = True
    End If
    If RedOre < 10 Or RedVespian < 40 Or RedBarracks = 0 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(7).Visible = False
    Else
    imgBuilding(7).Visible = True
    End If
    If RedOre < 5 Or RedVespian < 20 Or UnitType(ActiveUnit) <> "Worker" Then
    imgBuilding(6).Visible = False
    Else
    imgBuilding(6).Visible = True
    End If
    imgBuilding(8).Visible = False
    '''Worker
    If RedOre < 10 Or RedUnits = RedSupport Then
    imgBuildUnit(0).Visible = False
    Else
    imgBuildUnit(0).Visible = True
    End If
    '''Trooper
    If RedOre < 20 Or RedBarracks = 0 Or RedUnits = RedSupport Then
    imgBuildUnit(1).Visible = False
    Else
    imgBuildUnit(1).Visible = True
    End If
    '''Jedi
    If RedOre < 30 Or RedBarracks = 0 Or RedUnits = RedSupport Then
    imgBuildUnit(2).Visible = False
    Else
    imgBuildUnit(2).Visible = True
    End If
    '''Tank
    If RedOre < 40 Or RedMechLab = 0 Or RedUnits = RedSupport Then
    imgBuildUnit(3).Visible = False
    Else
    imgBuildUnit(3).Visible = True
    End If
    '''Gribble
    If RedOre < 80 Or RedHatchery = 0 Or RedUnits = RedSupport Then
    imgBuildUnit(4).Visible = False
    Else
    imgBuildUnit(4).Visible = True
    End If
    '''Pyromegalon
    If RedVespian < 18 Or RedGeothermal = 0 Or RedUnits = RedSupport Then
    imgBuildUnit(5).Visible = False
    Else
    imgBuildUnit(5).Visible = True
    End If
    imgBuildUnit(6).Visible = False
End If

End Sub

Private Sub TowerShoot(Unit)
Dim i
For i = 1 To 259
    If TileState(i) = "Defense Tower" And TileOwner(i) <> Turn And BuildingState(i) = 0 Then
    If imgTile(i).Top >= imgUnit(Unit).Top - 1440 And _
    imgTile(i).Top <= imgUnit(Unit).Top + 1440 Then
        If imgTile(i).Left >= imgUnit(Unit).Left - 1440 And _
        imgTile(i).Left <= imgUnit(Unit).Left + 1440 Then
        If Rnd < 2 / 3 Then Health(Unit) = Health(Unit) - 50
        linHealthBar.X2 = linHealthBar.X1 + 720 * (Health(ActiveUnit) / HP(ActiveUnit))
        linHealthBar.BorderColor = RGB(255 - (255 * (Health(ActiveUnit) / HP(ActiveUnit))), 255 * (Health(ActiveUnit) / HP(ActiveUnit)), 0)
        lblActiveUnit.Caption = UnitTeam(ActiveUnit) & " " & UnitType(ActiveUnit) & " HP " & Health(ActiveUnit) & "/" & HP(ActiveUnit) & " Moves: " & MovesLeft
        If Health(Unit) <= 0 Then
        Mode(Unit) = "Dead"
        If Health(Unit) < 0 Then Health(Unit) = 0
        imgUnit(Unit).Top = -720
        Call TurnCycle
        End If
    End If
    End If
    End If
Next i

End Sub

Private Sub BuildingAttack(Index As Integer)
Dim i
'''
'''===<<<IF THIS IS WORKING DON'T EVEN LOOK AT IT.  YOU'LL GET A MIGRAINE IN ABOUT 30 SECS>>>===
'''
    If imgTile(Index).Top >= imgUnit(ActiveUnit).Top - 720 * Range(ActiveUnit) And _
    imgTile(Index).Top <= imgUnit(ActiveUnit).Top + 720 * Range(ActiveUnit) Then
        If imgTile(Index).Left >= imgUnit(ActiveUnit).Left - 720 * Range(ActiveUnit) And _
        imgTile(Index).Left <= imgUnit(ActiveUnit).Left + 720 * Range(ActiveUnit) Then
            If TileOwner(Index) <> Turn Then
            '''Preliminaries
                If Rnd <= Accuracy(ActiveUnit) And BuildingState(Index) = 0 Then BuildingHealth(Index) = BuildingHealth(Index) - Firepower(ActiveUnit)
                '''if the building is to be destroyed
                    If BuildingHealth(Index) <= 0 Or BuildingState(Index) < 0 Then
                    BuildingHealth(Index) = 0
                    If TileState(Index) = "Barracks" Then
                        If Turn = "Blue" Then
                        'Kill Red's Barracks + Supported Buildings
                        If BuildingState(Index) = 0 Then RedBarracks = RedBarracks - 1
                                If RedBarracks = 0 Then
                                    For i = 0 To 259
                                        If TileState(i) = "Mech Lab" And TileOwner(i) <> Turn Then
                                        imgTile(i).Picture = imgLuna.Picture
                                        TileOwner(i) = ""
                                        TileState(i) = ""
                                        End If
                                    '''
                                        If TileState(i) = "Gribble Hatchery" And TileOwner(i) <> Turn Then
                                        imgTile(i).Picture = imgLuna.Picture
                                        TileOwner(i) = ""
                                        TileState(i) = ""
                                        End If
                                    Next i
                                RedMechLab = 0
                                RedHatchery = 0
                                End If
                        Else
                            'Kill Blue's Barracks etc.
                            If BuildingState(Index) = 0 Then BlueBarracks = BlueBarracks - 1
                                If BlueBarracks = 0 Then
                                    For i = 0 To 259
                                        If TileState(i) = "Mech Lab" And TileOwner(i) <> Turn Then
                                        imgTile(i).Picture = imgLuna.Picture
                                        TileOwner(i) = ""
                                        TileState(i) = ""
                                        End If
                                    '''
                                        If TileState(i) = "Gribble Hatchery" And TileOwner(i) <> Turn Then
                                        imgTile(i).Picture = imgLuna.Picture
                                        TileOwner(i) = ""
                                        TileState(i) = ""
                                        End If
                                    Next i
                                BlueMechLab = 0
                                BlueHatchery = 0
                                End If
                            End If
                        End If
                        If TileState(Index) = "Mech Lab" And BuildingState(Index) = 0 Then
                            If Turn = "Blue" Then
                            RedMechLab = RedMechLab - 1
                            Else
                            BlueMechLab = BlueMechLab - 1
                            End If
                        End If
                        If TileState(Index) = "Gribble Hatchery" And BuildingState(Index) = 0 Then
                            If Turn = "Blue" Then
                            RedHatchery = RedHatchery - 1
                            Else
                            BlueHatchery = BlueHatchery - 1
                            End If
                        End If
                        If TileState(Index) = "Geothermal Plant" And BuildingState(Index) = 0 Then
                        RedGeothermal = RedGeothermal - 1
                        End If
                        If TileState(Index) = "Cryogenics Lab" And BuildingState(Index) = 0 Then
                        BlueCryogenics = BlueCryogenics - 1
                        End If
                    TileState(Index) = ""
                    TileOwner(Index) = ""
                    imgTile(Index).Picture = imgLuna.Picture
                    '                On Error Resume Next
                    '                Call TowerShoot(ActiveUnit)
                    '            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
                    '''end destroy building
                    End If
                End If
            '''end preliminaries
            On Error Resume Next
            Call TowerShoot(ActiveUnit)
            If Mode(ActiveUnit) <> "Dead" Then Call TurnCycle
            End If
        '            If TileOwner(Index) <> Turn And BuildingState(Index) < 0 Then
        '            TileState(Index) = ""
        '            TileOwner(Index) = ""
        '            imgTile(Index).Picture = imgLuna.Picture
        End If
'    End If
End Sub

