VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Translator 2000"
   ClientHeight    =   3750
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   3120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK (OL)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1020
      TabIndex        =   2
      ToolTipText     =   "Click to translate"
      Top             =   1500
      Width           =   1035
   End
   Begin VB.TextBox txtEnter 
      Height          =   1000
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Type word(s) here"
      Top             =   420
      Width           =   3015
   End
   Begin VB.Label lblTranslate 
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   1000
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Translation appears here"
      Top             =   1980
      Width           =   3015
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File (Gime)"
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit (Ruiu)"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Position, InChr, OutChr As String
Dim InVal, OutVal

Private Sub cmdOK_Click()
lblTranslate.Caption = ""
For Position = 1 To Len(txtEnter.Text)
InChr = Mid(txtEnter.Text, Position, 1)
InVal = Asc(InChr)

Select Case InVal
Case Asc("A"), Asc("E"), Asc("I"), Asc("O"), Asc("U"), Asc("Y"), Asc(" "), Asc _
     ("a"), Asc("e"), Asc("i"), Asc("o"), Asc("u"), Asc("y"), Is _
     < 65, Is > 122 ' These are all non-letter characters
OutVal = InVal
Case Asc("Z"), Asc("z")
OutVal = Asc("b")
Case Asc("P"), Asc("p")
OutVal = Asc("q" & "u")
Case Else
OutVal = InVal + 1
End Select
OutChr = Chr(OutVal)
lblTranslate.Caption = lblTranslate.Caption & OutChr
Next Position
End Sub

Private Sub mnuQuit_Click()
End
End Sub

Private Sub txtEnter_Change()
'If Len(txtEnter.Text) >= 20 Then
'MsgBox "Please enter a word that makes sense.", 48, "Invalid Word"
'txtEnter.Text = ""
'End If
End Sub
