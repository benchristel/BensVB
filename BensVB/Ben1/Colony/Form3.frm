VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form3"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2415
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   300
      TabIndex        =   6
      Top             =   2040
      Width           =   1875
   End
   Begin VB.OptionButton optSouth 
      Caption         =   "South (No Points)"
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   1620
      Width           =   1755
   End
   Begin VB.OptionButton optMiddle 
      Caption         =   "Middle (+50 Points)"
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1755
   End
   Begin VB.OptionButton optNorth 
      Caption         =   "North (+100 Points)"
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1020
      Width           =   1755
   End
   Begin VB.TextBox txtColonyName 
      Height          =   315
      Left            =   420
      TabIndex        =   0
      Top             =   420
      Width           =   1635
   End
   Begin VB.Frame fraLocation 
      Caption         =   "Choose Your Location"
      Height          =   1155
      Left            =   300
      TabIndex        =   5
      Top             =   840
      Width           =   1875
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      Caption         =   "Enter the name of your colony."
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
If Len(txtColonyName.Text) < 15 Then
ColonyName = txtColonyName.Text
Form1.Show
Unload Form3
Else
MsgBox "Names are limeted to 15 characters", , "Error!"
End If
End Sub

Private Sub Form_Load()
Unload Form1
Randomize
RndName = Int(Rnd * 6)
Select Case RndName
Case Is = 0
RndName = "Plymouth"
Location = "North"
optNorth.Value = True
Case Is = 1
RndName = "Jamestown"
Location = "South"
optSouth.Value = True
Case Is = 2
RndName = "Philidelphia"
Location = "Middle"
optMiddle.Value = True
Case Is = 3
RndName = "New York"
Location = "Middle"
optMiddle.Value = True
Case Is = 4
RndName = "Williamsburg"
Location = "South"
optSouth.Value = True
Case Is = 5
RndName = "Boston"
Location = "North"
optNorth.Value = True
End Select
txtColonyName.Text = RndName
End Sub

Private Sub txtColonyName_Change()
If Len(txtColonyName.Text) > 15 Then
MsgBox "Names are limeted to 15 characters", , "Error!"
Exit Sub
End If
End Sub
