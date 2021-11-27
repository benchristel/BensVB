VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5460
   LinkTopic       =   "Form4"
   ScaleHeight     =   3195
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1980
      TabIndex        =   5
      Top             =   1860
      Width           =   1575
   End
   Begin VB.TextBox txtDay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3660
      TabIndex        =   4
      Text            =   "1"
      Top             =   780
      Width           =   915
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      Caption         =   "January"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2820
      TabIndex        =   3
      Top             =   60
      Width           =   2595
   End
   Begin VB.Label lblDeity 
      Caption         =   "Today's Deity: Electrum"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1020
      Width           =   2475
   End
   Begin VB.Label lblPower 
      Caption         =   "Today's Power: Heat"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   2475
   End
   Begin VB.Label lblElement 
      Caption         =   "Today's Element: Earth"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2475
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMode 
      Caption         =   "Mode"
      Begin VB.Menu mnuGtoC 
         Caption         =   "Gregorian to Ceoete"
      End
      Begin VB.Menu mnuCtoG 
         Caption         =   "Ceoete to Gregorian"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Month, Day, DayOfYear, Element, Power, God

Private Sub cmdGo_Click()
Select Case Month
Case Is = "January"
DayOfYear = 0
Case Is = "February"
DayOfYear = 31
Case Is = "March"
DayOfYear = 59
Case Is = "April"
DayOfYear = 90
Case Is = "May"
DayOfYear = 120
Case Is = "June"
DayOfYear = 151
Case Is = "July"
DayOfYear = 181
Case Is = "August"
DayOfYear = 212
Case Is = "September"
DayOfYear = 243
Case Is = "October"
DayOfYear = 273
Case Is = "November"
DayOfYear = 304
Case Is = "December"
DayOfYear = 334
End Select
DayOfYear = DayOfYear + Day
Select Case DayOfYear Mod 4
Case Is = 1
Element = "Earth"
Case Is = 2
Element = "Water"
Case Is = 3
Element = "Air"
Case Is = 0
Element = "Fire"
End Select
Select Case DayOfYear Mod 7
Case Is = 1
Power = "Heat"
Case Is = 2
Power = "Light"
Case Is = 3
Power = "Gravity"
Case Is = 4
Power = "Magnetism"
Case Is = 5
Power = "Sound"
Case Is = 6
Power = "Wind"
Case Is = 0
Power = "Life"
End Select
Select Case DayOfYear Mod 13
Case Is = 1
God = "Electrum"
Case Is = 2
God = "Batlax"
Case Is = 3
God = "Paxor"
Case Is = 4
God = "Atlantis"
Case Is = 5
God = "Soleax"
Case Is = 6
God = "Lunalor"
Case Is = 7
God = "Lyra"
Case Is = 8
God = "Morz"
Case Is = 9
God = "Arctur"
Case Is = 10
God = "Kumbolo"
Case Is = 11
God = "Celidor"
Case Is = 12
God = "Viznor"
Case Is = 0
God = "Axtrax"
End Select
lblElement.Caption = "Today's Element: " & Element
lblPower.Caption = "Today's Power: " & Power
lblDeity.Caption = "Today's Deity: " & God
If Month = "December" And Day = 31 Then
lblElement.Caption = "Today's Element: N/A"
lblPower.Caption = "Today's Power: N/A"
lblDeity.Caption = "Today's Deity: N/A"
End If
End Sub

Private Sub Form_Load()
Month = "January"
Day = 1
End Sub

Private Sub lblMonth_Click()
Select Case Month
Case Is = "January"
Month = "February"
Case Is = "February"
Month = "March"
Case Is = "March"
Month = "April"
Case Is = "April"
Month = "May"
Case Is = "May"
Month = "June"
Case Is = "June"
Month = "July"
Case Is = "July"
Month = "August"
Case Is = "August"
Month = "September"
Case Is = "September"
Month = "October"
Case Is = "October"
Month = "November"
Case Is = "November"
Month = "December"
Case Is = "December"
Month = "January"
End Select
lblMonth.Caption = Month
End Sub

Private Sub txtDay_Change()
Day = Int(Val(txtDay.Text))
End Sub
