VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "Start"
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox picScreen 
      Height          =   6375
      Left            =   120
      ScaleHeight     =   421
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   677
      TabIndex        =   11
      Top             =   1200
      Width           =   10215
   End
   Begin VB.TextBox txtFadeSpeed 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "30"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtBlue 
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   9
      Text            =   "0"
      ToolTipText     =   "Blue"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtBlue 
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   8
      Text            =   "0"
      ToolTipText     =   "Blue"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox txtGreen 
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "Green"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtGreen 
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   6
      Text            =   "0"
      ToolTipText     =   "Green"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtRed 
      Height          =   285
      Index           =   1
      Left            =   4800
      TabIndex        =   5
      Text            =   "0"
      ToolTipText     =   "Red"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtRed 
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Red"
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00000000&
      Height          =   975
      Index           =   1
      Left            =   3720
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H00000000&
      Height          =   975
      Index           =   0
      Left            =   1440
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtTone 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "440"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtFrequency 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "10"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Frequency As Single, Period As Double, Tone As Integer
Dim State As Integer, Stopped As Boolean
Dim Red(0 To 1) As Integer, Green(0 To 1) As Integer, Blue(0 To 1) As Integer
Dim StrobeDC(0 To 1) As Long
Const KEY_PRESSED As Integer = &H1000

Private Sub cmdGo_Click()
If Stopped = True Then
Stopped = False
StretchBlt StrobeDC(0), 0, 0, 900, 900, picColor(0).hdc, 0, 0, 60, 60, vbSrcCopy
StretchBlt StrobeDC(1), 0, 0, 900, 900, picColor(1).hdc, 0, 0, 60, 60, vbSrcCopy
Call MainLoop
Else
Stopped = True
End If
End Sub

Private Sub Form_Load()
StrobeDC(0) = GenerateDC(App.Path & "\a.bmp")
StrobeDC(1) = GenerateDC(App.Path & "\a.bmp")
Frequency = 10
Tone = 440
Stopped = True
End Sub

Public Sub MainLoop()
Dim CurrentTick As Long, LastTick As Long, Residue As Double
Period = 500 / Frequency
'txtRed(0).Text = picColor(1).hdc
LastTick = GetTickCount
Do While Stopped = False
CurrentTick = GetTickCount
'If LastTick = 0 Then GoTo firstframe
'    If CurrentTick = LastTick Then GoTo loopthis

If CurrentTick - LastTick >= Period Then
'Residue = Residue - Period + CurrentTick - LastTick
If State = 1 Then
State = 0
Else
State = 1
'Beep Tone, Period / 2
End If
'MsgBox picColor(State).hdc
'picScreen.BackColor = picColor(State).BackColor
If State = 1 Then
BitBlt picScreen.hdc, 0, 0, 300, 300, StrobeDC(1), 0, 0, vbSrcCopy
Else
BitBlt picScreen.hdc, 0, 0, 300, 300, StrobeDC(0), 0, 0, vbSrcCopy
End If

LastTick = CurrentTick
End If

        
        
        'Sleep 2
    DoEvents
firstframe:
loopthis:
Loop
End Sub

Private Sub Form_Resize()
picScreen.Width = Form1.Width
picScreen.Height = Form1.Height - 1200
End Sub

Private Sub txtFrequency_Change()
If Val(txtFrequency.Text) >= 1 And Val(txtFrequency.Text) < 30 Then
Frequency = txtFrequency.Text
End If
End Sub
Private Sub txttone_Change()
If Val(txtTone.Text) >= 20 And Val(txtTone.Text) < 30000 Then
Tone = txtTone.Text
End If
End Sub
Private Sub txtRed_Change(Index As Integer)
If Val(txtRed(Index).Text) >= 0 And Val(txtRed(Index).Text) < 256 Then
    Red(Index) = Int(Val(txtRed(Index).Text))
    UpdateDisplay
End If
End Sub
Private Sub txtGreen_Change(Index As Integer)
If Val(txtGreen(Index).Text) >= 0 And Val(txtGreen(Index).Text) < 256 Then
    Green(Index) = Int(Val(txtGreen(Index).Text))
    UpdateDisplay
End If
End Sub
Private Sub txtBlue_Change(Index As Integer)
If Val(txtBlue(Index).Text) >= 0 And Val(txtBlue(Index).Text) < 256 Then
    Blue(Index) = Int(Val(txtBlue(Index).Text))
    UpdateDisplay
End If
End Sub

Public Sub UpdateDisplay()
Dim i
For i = 0 To 1
    picColor(i).BackColor = RGB(Red(i), Green(i), Blue(i))
Next i
End Sub

