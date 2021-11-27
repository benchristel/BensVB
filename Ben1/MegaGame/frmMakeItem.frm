VERSION 5.00
Begin VB.Form frmMakeItem 
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   LinkTopic       =   "Form2"
   ScaleHeight     =   3765
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cook Anchovies"
      Height          =   315
      Index           =   15
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   300
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cook Shrimps"
      Height          =   315
      Index           =   14
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   0
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fish for Fish"
      Height          =   315
      Index           =   13
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   300
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fill Bucket"
      Height          =   315
      Index           =   2
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Make Bronze Axe"
      Height          =   315
      Index           =   11
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Make Tinderbox"
      Height          =   315
      Index           =   9
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   300
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Make Copper Hammer"
      Height          =   315
      Index           =   6
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Carve Stone Anvil"
      Height          =   315
      Index           =   7
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   300
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Build Furnace"
      Height          =   315
      Index           =   4
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Make Fishing Net"
      Height          =   315
      Index           =   12
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   900
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Smelt Bronze"
      Height          =   315
      Index           =   10
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Smelt Tin"
      Height          =   315
      Index           =   8
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   300
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Smelt Copper"
      Height          =   315
      Index           =   5
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Make Brick"
      Height          =   315
      Index           =   3
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Make Bucket"
      Height          =   315
      Index           =   1
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   300
      Width           =   1755
   End
   Begin VB.CommandButton cmdMakeItem 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Make Flint Axe"
      Height          =   315
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1755
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   435
      Left            =   7560
      TabIndex        =   0
      Top             =   3300
      Width           =   1215
   End
End
Attribute VB_Name = "frmMakeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
Unload Me
End Sub

Private Sub cmdMakeItem_Click(Index As Integer)
Select Case Index
Case Is = 0 ' flint axe
Call MakeFlintAxe
Case Is = 1 ' bucket
Call MakeBucket
Case Is = 2 'fill bucket with water
Call FillBucket
Case Is = 3 'make clay into bricks using water
Call MakeBrick
Case Is = 4 'build a furnace
Call BuildFurnace
Case Is = 5 'smelt copper
Call SmeltCopper
Case Is = 6 'make copper hammer
Call MakeCopperHammer
Case Is = 7 'make stone anvil
Call MakeStoneAnvil
Case Is = 8 'smelt tin
Call SmeltTin
Case Is = 9 'make tinderbox
Call MakeTinderbox
Case Is = 10 'smelt bronze
Call SmeltBronze
Case Is = 11 'make bronze axe
Call MakeBronzeAxe
Case Is = 12 'make fishing net
Call MakeFishingNet
Case Is = 13 'fish for fish
If Fish > 0 Then Call FishForFish
Case Is = 14 'cook shrimps
Call CookShrimps
End Select
End Sub

