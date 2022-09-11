VERSION 5.00
Begin VB.Form frmDesigner 
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbTails 
      Height          =   2715
      ItemData        =   "frmDesigner.frx":0000
      Left            =   4920
      List            =   "frmDesigner.frx":0085
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Text            =   "Tailplanes"
      Top             =   60
      Width           =   1575
   End
   Begin VB.ComboBox cbWings 
      Height          =   2715
      ItemData        =   "frmDesigner.frx":0276
      Left            =   3300
      List            =   "frmDesigner.frx":02FE
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Text            =   "Wings"
      Top             =   60
      Width           =   1575
   End
   Begin VB.ComboBox cbFuse 
      Height          =   2715
      ItemData        =   "frmDesigner.frx":04CF
      Left            =   1680
      List            =   "frmDesigner.frx":0536
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Text            =   "Fuselages"
      Top             =   60
      Width           =   1575
   End
   Begin VB.ComboBox cbNose 
      Height          =   2715
      ItemData        =   "frmDesigner.frx":06A9
      Left            =   60
      List            =   "frmDesigner.frx":0707
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Text            =   "Noses"
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "frmDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

