VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   10530
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInventoryDown 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Page Down"
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10080
      Width           =   3495
   End
   Begin VB.CommandButton cmdInventoryUp 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Page Up"
      Height          =   375
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5160
      Width           =   3495
   End
   Begin VB.CommandButton cmdBuy 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buy"
      Height          =   495
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSellAll 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sell All"
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdSell 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sell"
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Feet: [Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8280
      TabIndex        =   27
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label lblGloves 
      BackColor       =   &H00000000&
      Caption         =   "Gloves: [Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8280
      TabIndex        =   26
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label lblAmulet 
      BackColor       =   &H00000000&
      Caption         =   "Amulet: [Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8280
      TabIndex        =   25
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   11
      Left            =   8280
      TabIndex        =   24
      Top             =   9600
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   10
      Left            =   8280
      TabIndex        =   23
      Top             =   9240
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   9
      Left            =   8280
      TabIndex        =   22
      Top             =   8880
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   8
      Left            =   8280
      TabIndex        =   21
      Top             =   8520
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   7
      Left            =   8280
      TabIndex        =   20
      Top             =   8160
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   6
      Left            =   8280
      TabIndex        =   19
      Top             =   7800
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   5
      Left            =   8280
      TabIndex        =   18
      Top             =   7440
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   4
      Left            =   8280
      TabIndex        =   17
      Top             =   7080
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   3
      Left            =   8280
      TabIndex        =   16
      Top             =   6720
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   2
      Left            =   8280
      TabIndex        =   15
      Top             =   6360
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   8280
      TabIndex        =   14
      Top             =   6000
      Width           =   3495
   End
   Begin VB.Label lblInventory 
      BackColor       =   &H00000000&
      Caption         =   "[Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   8280
      TabIndex        =   13
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Label lblLinkToLocation 
      BackColor       =   &H00808080&
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   5640
      Width           =   3495
   End
   Begin VB.Label lblMessages 
      BackColor       =   &H00000000&
      Caption         =   "Welcome to Mithgoldh."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   8280
      TabIndex        =   9
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label lblItem 
      BackColor       =   &H00000000&
      Caption         =   "Wooden Sword"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblLegs 
      BackColor       =   &H00000000&
      Caption         =   "Legs: [Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label lblBody 
      BackColor       =   &H00000000&
      Caption         =   "Body: [Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label lblLeftHand 
      BackColor       =   &H00000000&
      Caption         =   "Left Hand: [Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lblRightHand 
      BackColor       =   &H00000000&
      Caption         =   "Right Hand: [Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lblHead 
      BackColor       =   &H00000000&
      Caption         =   "Head: [Empty]"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8280
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
