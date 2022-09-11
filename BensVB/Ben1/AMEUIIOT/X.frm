VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H80000009&
      Height          =   675
      Left            =   8700
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   60
      Width           =   1755
   End
   Begin VB.CommandButton cmdGo 
      Height          =   675
      Left            =   6900
      TabIndex        =   0
      Top             =   60
      Width           =   1755
   End
   Begin VB.Label lblSelect 
      BackColor       =   &H00FFFFFF&
      Height          =   675
      Index           =   7
      Left            =   6120
      TabIndex        =   9
      Top             =   5100
      Width           =   675
   End
   Begin VB.Label lblSelect 
      BackColor       =   &H000000FF&
      Height          =   675
      Index           =   6
      Left            =   6120
      TabIndex        =   8
      Top             =   4380
      Width           =   675
   End
   Begin VB.Label lblSelect 
      BackColor       =   &H000080FF&
      Height          =   675
      Index           =   5
      Left            =   6120
      TabIndex        =   7
      Top             =   3660
      Width           =   675
   End
   Begin VB.Label lblSelect 
      BackColor       =   &H0000FFFF&
      Height          =   675
      Index           =   4
      Left            =   6120
      TabIndex        =   6
      Top             =   2940
      Width           =   675
   End
   Begin VB.Label lblSelect 
      BackColor       =   &H0000C000&
      Height          =   675
      Index           =   3
      Left            =   6120
      TabIndex        =   5
      Top             =   2220
      Width           =   675
   End
   Begin VB.Label lblSelect 
      BackColor       =   &H00FF0000&
      Height          =   675
      Index           =   2
      Left            =   6120
      TabIndex        =   4
      Top             =   1500
      Width           =   675
   End
   Begin VB.Label lblSelect 
      BackColor       =   &H00C000C0&
      Height          =   675
      Index           =   1
      Left            =   6120
      TabIndex        =   3
      Top             =   780
      Width           =   675
   End
   Begin VB.Label lblSelect 
      BackColor       =   &H00000000&
      Height          =   675
      Index           =   0
      Left            =   6120
      TabIndex        =   2
      Top             =   60
      Width           =   675
   End
   Begin VB.Label lblDisplay 
      Height          =   4995
      Left            =   6900
      TabIndex        =   1
      Top             =   780
      Width           =   3555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   10
      X1              =   5940
      X2              =   5940
      Y1              =   60
      Y2              =   5760
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   63
      Left            =   5100
      Top             =   5100
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   62
      Left            =   4380
      Top             =   5100
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   61
      Left            =   3660
      Top             =   5100
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   60
      Left            =   2940
      Top             =   5100
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   59
      Left            =   2220
      Top             =   5100
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   58
      Left            =   1500
      Top             =   5100
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   57
      Left            =   780
      Top             =   5100
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   56
      Left            =   60
      Top             =   5100
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   55
      Left            =   5100
      Top             =   4380
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   54
      Left            =   4380
      Top             =   4380
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   53
      Left            =   3660
      Top             =   4380
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   52
      Left            =   2940
      Top             =   4380
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   51
      Left            =   2220
      Top             =   4380
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   50
      Left            =   1500
      Top             =   4380
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   49
      Left            =   780
      Top             =   4380
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   48
      Left            =   60
      Top             =   4380
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   47
      Left            =   5100
      Top             =   3660
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   46
      Left            =   4380
      Top             =   3660
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   45
      Left            =   3660
      Top             =   3660
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   44
      Left            =   2940
      Top             =   3660
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   43
      Left            =   2220
      Top             =   3660
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   42
      Left            =   1500
      Top             =   3660
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   41
      Left            =   780
      Top             =   3660
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   40
      Left            =   60
      Top             =   3660
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   39
      Left            =   5100
      Top             =   2940
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   38
      Left            =   4380
      Top             =   2940
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   37
      Left            =   3660
      Top             =   2940
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   36
      Left            =   2940
      Top             =   2940
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   35
      Left            =   2220
      Top             =   2940
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   34
      Left            =   1500
      Top             =   2940
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   33
      Left            =   780
      Top             =   2940
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   32
      Left            =   60
      Top             =   2940
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   31
      Left            =   5100
      Top             =   2220
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   30
      Left            =   4380
      Top             =   2220
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   29
      Left            =   3660
      Top             =   2220
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   28
      Left            =   2940
      Top             =   2220
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   27
      Left            =   2220
      Top             =   2220
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   26
      Left            =   1500
      Top             =   2220
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   25
      Left            =   780
      Top             =   2220
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   24
      Left            =   60
      Top             =   2220
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   23
      Left            =   5100
      Top             =   1500
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   22
      Left            =   4380
      Top             =   1500
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   21
      Left            =   3660
      Top             =   1500
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   20
      Left            =   2940
      Top             =   1500
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   19
      Left            =   2220
      Top             =   1500
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   18
      Left            =   1500
      Top             =   1500
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   17
      Left            =   780
      Top             =   1500
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   16
      Left            =   60
      Top             =   1500
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   15
      Left            =   5100
      Top             =   780
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   14
      Left            =   4380
      Top             =   780
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   13
      Left            =   3660
      Top             =   780
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   12
      Left            =   2940
      Top             =   780
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   11
      Left            =   2220
      Top             =   780
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   10
      Left            =   1500
      Top             =   780
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   9
      Left            =   780
      Top             =   780
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   8
      Left            =   60
      Top             =   780
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   7
      Left            =   5100
      Top             =   60
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   6
      Left            =   4380
      Top             =   60
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   5
      Left            =   3660
      Top             =   60
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   4
      Left            =   2940
      Top             =   60
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   3
      Left            =   2220
      Top             =   60
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   2
      Left            =   1500
      Top             =   60
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   1
      Left            =   780
      Top             =   60
      Width           =   675
   End
   Begin VB.Shape shpDirector 
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   675
      Index           =   0
      Left            =   60
      Top             =   60
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Random

Private Sub cmdClear_Click()
lblDisplay.Caption = ""
End Sub

Private Sub cmdGo_Click()
Dim Step, i, Present
Present = Left(lblDisplay.Caption, 1)
Step = 1
Do Until Step = Len(lblDisplay.Caption)
    For i = 0 To 63
    If Step Mod 2 = 0 Then
        If shpDirector(i).Top = shpDirector(Present).Top And _
                shpDirector(i).FillColor = shpDirector(Mid(lblDisplay.Caption, Step + 1, 1)).FillColor Then
        Present = i
        Exit For
        End If
    Else
        If shpDirector(i).Left = shpDirector(Present).Left And _
                shpDirector(i).FillColor = shpDirector(Mid(lblDisplay.Caption, Step + 1, 1)).FillColor Then
        Present = i
        Exit For
        End If
    End If
    Next i
    Step = Step + 1
Loop



    
    





End Sub

Private Sub lblSelect_Click(Index As Integer)
lblDisplay.Caption = lblDisplay.Caption & Index
End Sub
