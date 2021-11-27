VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   17
      Top             =   1260
      Width           =   2295
   End
   Begin VB.CommandButton cmdDie3 
      Height          =   735
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdDie2 
      Height          =   735
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdDie1 
      Height          =   735
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblEnter 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   2400
      TabIndex        =   18
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   13
      Left            =   10200
      TabIndex        =   13
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   12
      Left            =   9420
      TabIndex        =   12
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   11
      Left            =   8640
      TabIndex        =   11
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   10
      Left            =   7860
      TabIndex        =   10
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   9
      Left            =   7080
      TabIndex        =   9
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   8
      Left            =   6300
      TabIndex        =   8
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   7
      Left            =   5520
      TabIndex        =   7
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   4740
      TabIndex        =   6
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   3960
      TabIndex        =   5
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   3180
      TabIndex        =   4
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   3
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   1620
      TabIndex        =   2
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   60
      Width           =   735
   End
   Begin VB.Label lblSpace 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Num1, Num2, Num3

Private Sub cmdDie1_Click()
Dim i, x
lblEnter.Caption = lblEnter.Caption & cmdDie1.Caption
cmdDie1.BackColor = &H8000000F
cmdDie1.Caption = ""
If Len(lblEnter.Caption) = 3 Then
For i = 0 To 13
If lblSpace(i).Caption = "" Then lblSpace(i).Enabled = True
Next i
Check
End If
'For i = 0 To 13
'If lblSpace(i).Caption = "" Then
'For x = 0 To i
'If lblSpace(x).Caption = "" Or lblSpace(x).Caption < lblSpace(i).Caption Then Exit For
'Next x
'Next i
End Sub

Private Sub cmdDie2_Click()
Dim i
lblEnter.Caption = lblEnter.Caption & cmdDie2.Caption
cmdDie2.BackColor = &H8000000F
cmdDie2.Caption = ""
'For i = 0 To 13
'lblSpace(i).Enabled = True
'Next i
If Len(lblEnter.Caption) = 3 Then
For i = 0 To 13
If lblSpace(i).Caption = "" Then lblSpace(i).Enabled = True
Next i
Check
End If
End Sub

Private Sub cmdDie3_Click()
Dim i
lblEnter.Caption = lblEnter.Caption & cmdDie3.Caption
cmdDie3.BackColor = &H8000000F
cmdDie3.Caption = ""
If Len(lblEnter.Caption) = 3 Then
For i = 0 To 13
If lblSpace(i).Caption = "" Then lblSpace(i).Enabled = True
Next i
Check
End If
'For i = 0 To 13
'lblSpace(i).Enabled = True
'Next i
End Sub

Private Sub cmdRoll_Click()
Dim i
Randomize
lblEnter.Caption = ""
For i = 0 To 13
    lblSpace(i).Enabled = False
Next i
Num1 = Int(Rnd * 6 + 1)
Num2 = Int(Rnd * 6 + 1)
Num3 = Int(Rnd * 6 + 4)
With cmdDie1
    .Caption = Num1
    .BackColor = &HFFFF00
End With
With cmdDie2
    .Caption = Num2
    .BackColor = &HFFFF00
End With
With cmdDie3
    .Caption = Num3
    .BackColor = &HFFFF00
End With
End Sub



Private Sub lblSpace_Click(Index As Integer)
Dim i
    lblSpace(Index).Caption = lblEnter.Caption
    For i = 0 To 13
    lblSpace(i).Enabled = False
Next i
End Sub

Private Sub Check()
Dim i, x, Disabled(0 To 13)
For i = 0 To 13
For x = 0 To i
If lblSpace(x).Caption < lblEnter.Caption Then
lblSpace(i).Enabled = True
Else
lblSpace(i).Enabled = False
Disabled(i) = True
Exit For
End If
Next x
For x = 0 To 13
If lblSpace(x).Caption > lblEnter.Caption Or lblEnter.Caption = "" And Disabled(i) <> True Then
lblSpace(i).Enabled = True
Else
lblSpace(i).Enabled = False
Exit For
End If
Next x
Next i
End Sub
