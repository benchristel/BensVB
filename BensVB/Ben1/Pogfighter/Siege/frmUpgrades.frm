VERSION 5.00
Begin VB.Form frmUpgrade 
   BackColor       =   &H00000000&
   Caption         =   "Upgrade Station"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmUpgrades.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShootBox 
      Caption         =   "Rangefinder $15000"
      Height          =   495
      Left            =   1620
      TabIndex        =   15
      Top             =   2160
      Width           =   1515
   End
   Begin VB.CommandButton cmdStunGun 
      Caption         =   "Stun Gun $10000"
      Height          =   495
      Left            =   60
      TabIndex        =   14
      Top             =   2700
      Width           =   1515
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3720
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000C000&
      Caption         =   "Upgrade"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3720
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Directed Blaster $6500"
      Height          =   495
      Left            =   3180
      TabIndex        =   9
      Top             =   1620
      Width           =   1455
   End
   Begin VB.CommandButton cmdMine1 
      Caption         =   "High Energy Bomb $3000"
      Height          =   495
      Left            =   1620
      TabIndex        =   8
      Top             =   1620
      Width           =   1515
   End
   Begin VB.CommandButton cmdMissile2 
      Caption         =   "Neutron Missile $4000"
      Height          =   495
      Left            =   3180
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdMissile1 
      Caption         =   "Proton Torpedo $1000"
      Height          =   495
      Left            =   1620
      TabIndex        =   6
      Top             =   1080
      Width           =   1515
   End
   Begin VB.CommandButton cmdRadar 
      Caption         =   "Radar $15000"
      Height          =   495
      Left            =   60
      TabIndex        =   4
      Top             =   2160
      Width           =   1515
   End
   Begin VB.CommandButton cmdMineD 
      Caption         =   "Mine Dropper $8500"
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   1620
      Width           =   1515
   End
   Begin VB.CommandButton cmdMissileL 
      Caption         =   "Missile Launcher $5000"
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   1080
      Width           =   1515
   End
   Begin VB.CommandButton cmdExtraLife 
      Caption         =   "Extra Life $10000"
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   1515
   End
   Begin VB.Label lblGrandTotal 
      BackColor       =   &H00000000&
      Caption         =   "Total: $0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   60
      TabIndex        =   11
      Top             =   3780
      Width           =   1515
   End
   Begin VB.Label lblTotal2 
      BackColor       =   &H00000000&
      Caption         =   "Total: $0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2400
      TabIndex        =   10
      Top             =   3240
      Width           =   1515
   End
   Begin VB.Label lblTotal1 
      BackColor       =   &H00000000&
      Caption         =   "Total: $0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   60
      TabIndex        =   5
      Top             =   3240
      Width           =   1515
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H0000C000&
      Caption         =   "Upgrade: Anadon Fighter"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   4575
   End
End
Attribute VB_Name = "frmUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
frmPogfighter.tmrTime.Enabled = True
Unload frmUpgrade
End Sub

Private Sub Form_Load()
frmUpgrade.Visible = True
frmUpgrade.ZOrder
End Sub
