VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Fractal Generator"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   ScaleHeight     =   14850
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBlue 
      Height          =   285
      Left            =   6360
      TabIndex        =   20
      Text            =   "255"
      Top             =   13440
      Width           =   495
   End
   Begin VB.TextBox txtGreen 
      Height          =   285
      Left            =   5880
      TabIndex        =   19
      Text            =   "255"
      Top             =   13440
      Width           =   495
   End
   Begin VB.TextBox txtRed 
      Height          =   285
      Left            =   5400
      TabIndex        =   18
      Text            =   "255"
      Top             =   13440
      Width           =   495
   End
   Begin VB.TextBox txtRotate 
      Height          =   285
      Left            =   5400
      TabIndex        =   17
      Text            =   "0"
      Top             =   13200
      Width           =   1455
   End
   Begin VB.TextBox txtGenScale 
      Height          =   285
      Left            =   5400
      TabIndex        =   14
      Text            =   "1"
      Top             =   12960
      Width           =   1455
   End
   Begin VB.TextBox txtIterations 
      Height          =   285
      Left            =   5400
      TabIndex        =   12
      Text            =   "5"
      Top             =   12720
      Width           =   1455
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Generate Fractal"
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   12720
      Width           =   1815
   End
   Begin VB.TextBox txtScale 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   13680
      Width           =   1455
   End
   Begin VB.TextBox txtTheta 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   13440
      Width           =   1455
   End
   Begin VB.TextBox txtRCoord 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   13200
      Width           =   1455
   End
   Begin VB.TextBox txtYCoord 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   12960
      Width           =   1455
   End
   Begin VB.TextBox txtXCoord 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   12720
      Width           =   1455
   End
   Begin VB.PictureBox picScreen 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   12495
      Left            =   120
      ScaleHeight     =   829
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1253
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   18855
   End
   Begin VB.Label lblColor 
      Alignment       =   1  'Right Justify
      Caption         =   "Color:"
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   13440
      Width           =   855
   End
   Begin VB.Label lblRotate 
      Alignment       =   1  'Right Justify
      Caption         =   "Rotate:"
      Height          =   255
      Left            =   4440
      TabIndex        =   16
      Top             =   13200
      Width           =   855
   End
   Begin VB.Label lblGenScale 
      Alignment       =   1  'Right Justify
      Caption         =   "Scale:"
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   12960
      Width           =   855
   End
   Begin VB.Label lblIterations 
      Alignment       =   1  'Right Justify
      Caption         =   "Iterations:"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   12720
      Width           =   855
   End
   Begin VB.Label lblScale 
      Alignment       =   1  'Right Justify
      Caption         =   "Scale:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   13680
      Width           =   750
   End
   Begin VB.Label lblR 
      Alignment       =   1  'Right Justify
      Caption         =   "R:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   13200
      Width           =   750
   End
   Begin VB.Label lblTheta 
      Alignment       =   1  'Right Justify
      Caption         =   "Theta:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   13440
      Width           =   750
   End
   Begin VB.Label lblYCoord 
      Alignment       =   1  'Right Justify
      Caption         =   "Y:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   12960
      Width           =   750
   End
   Begin VB.Label lblXCoord 
      Alignment       =   1  'Right Justify
      Caption         =   "X:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   12720
      Width           =   750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DesignMode As Boolean
Dim TemplatePoint() As Point
Dim TemplatePointCount As Integer
Dim ActivePoint As Integer
Dim CenterX As Double, CenterY As Double
Dim OldPoint() As Point, NewPoint() As Point, OldCount As Integer, NewCount As Integer
Const Pi = 3.14159265358979

Private Sub cmdGo_Click()
If TemplatePointCount > 0 And DesignMode = True Then
DesignMode = False
Else
If DesignMode = True Then MsgBox "Add fractal template points before generating a fractal.", vbOKOnly, "Error!"
End If
End Sub

Private Sub Form_Load()

CenterX = picScreen.ScaleWidth / 2
CenterY = picScreen.ScaleHeight / 2
DesignMode = True
Call RedrawScreen
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim n As Integer, i As Integer, k As Integer, startx As Single, starty As Single
Dim genscale As Double, rotate As Double, iterations As Integer, colorred As Integer, colorgreen As Integer, colorblue As Integer
If DesignMode = True Then
Select Case Button
Case Is = 1
    For i = 1 To TemplatePointCount
    If Distance((X), (Y), TemplatePoint(i).XCoord + CenterX, TemplatePoint(i).YCoord + CenterY) <= 7 Then
    ActivePoint = i
    Call RedrawScreen
    End If
    Next i
Case Is = 2
    TemplatePointCount = TemplatePointCount + 1
    ReDim Preserve TemplatePoint(1 To TemplatePointCount)
    With TemplatePoint(TemplatePointCount)
    .RCoord = Distance((X), (Y), CenterX, CenterY)
    .ScaleFactor = TemplatePoint(TemplatePointCount).RCoord / 100
    .Theta = Atn2(-(Y - CenterY), (X - CenterX))
    .XCoord = X - CenterX
    .YCoord = Y - CenterY
    End With
    ActivePoint = TemplatePointCount
    Call RedrawScreen
End Select
Else 'designmode= false
On Error GoTo endit
picScreen.Cls
startx = X
starty = Y
genscale = Val(txtGenScale.Text)
rotate = Val(txtRotate.Text)
iterations = Val(txtIterations.Text)
colorred = Val(txtRed.Text)
colorgreen = Val(txtGreen.Text)
colorblue = Val(txtBlue.Text)
OldCount = 1
ReDim OldPoint(1 To 1)
With OldPoint(1)
    .ScaleFactor = genscale
    .XCoord = startx
    .YCoord = starty
    .Theta = rotate
End With
ReDim NewPoint(1 To TemplatePointCount)
picScreen.Circle (startx, starty), 2, RGB(colorred, colorgreen, colorblue)
For n = 1 To iterations
ReDim NewPoint(1 To TemplatePointCount ^ n)
For i = 1 To OldCount
For k = 1 To TemplatePointCount
With NewPoint(TemplatePointCount * (i - 1) + k)
    .Theta = TemplatePoint(k).Theta + OldPoint(i).Theta
    .RCoord = TemplatePoint(k).RCoord * OldPoint(i).ScaleFactor
    .ScaleFactor = TemplatePoint(k).ScaleFactor * OldPoint(i).ScaleFactor
    .XCoord = OldPoint(i).XCoord + NewPoint(TemplatePointCount * (i - 1) + k).RCoord * Sin(NewPoint(TemplatePointCount * (i - 1) + k).Theta * Pi / 180)
    .YCoord = OldPoint(i).YCoord + NewPoint(TemplatePointCount * (i - 1) + k).RCoord * -Cos(NewPoint(TemplatePointCount * (i - 1) + k).Theta * Pi / 180)
End With
picScreen.Circle (NewPoint(TemplatePointCount * (i - 1) + k).XCoord, NewPoint(TemplatePointCount * (i - 1) + k).YCoord), 2, RGB(colorred, colorgreen, colorblue)
Next k
Next i
OldCount = TemplatePointCount ^ n
ReDim OldPoint(1 To OldCount)
For i = 1 To OldCount
OldPoint(i) = NewPoint(i)
Next i
Next n
End If
endit:
End Sub

Public Sub RedrawScreen()
picScreen.Cls
Dim i As Integer
    picScreen.Circle (CenterX, CenterY), 4, RGB(255, 0, 0)
    picScreen.Circle (CenterX, CenterY), 100, RGB(100, 0, 0)
For i = 1 To TemplatePointCount
If i = ActivePoint Then
    picScreen.Circle (TemplatePoint(i).XCoord + CenterX, TemplatePoint(i).YCoord + CenterY), 4, RGB(0, 255, 0)
    picScreen.Circle (TemplatePoint(i).XCoord + CenterX, TemplatePoint(i).YCoord + CenterY), 100 * TemplatePoint(i).ScaleFactor, RGB(0, 100, 0)
Else
    picScreen.Circle (TemplatePoint(i).XCoord + CenterX, TemplatePoint(i).YCoord + CenterY), 4, RGB(0, 0, 255)
    picScreen.Circle (TemplatePoint(i).XCoord + CenterX, TemplatePoint(i).YCoord + CenterY), 100 * TemplatePoint(i).ScaleFactor, RGB(0, 0, 100)
End If
Next i
If ActivePoint > 0 Then
txtXCoord.Text = TemplatePoint(ActivePoint).XCoord
txtYCoord.Text = TemplatePoint(ActivePoint).YCoord
txtRCoord.Text = TemplatePoint(ActivePoint).RCoord
txtTheta.Text = TemplatePoint(ActivePoint).Theta
txtScale.Text = TemplatePoint(ActivePoint).ScaleFactor
End If
End Sub
Public Function Distance(X1 As Double, Y1 As Double, X2 As Double, Y2 As Double) As Double
Distance = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
End Function

Private Function Atn2(X As Double, Y As Double) As Double

 'The Atn2 function takes the two sides of a right
'triangle (as double) and returns the corresponding angle
'in radians (as double). x is the length of the side adjacent
'to the angle and y is the length of the side opposite
'the angle.
'The range of the result is 0 to 2pi radians.
'To convert degrees to radians, multiply degrees by pi/180.
'To convert radians to degrees, multiply radians by 180/pi.


Select Case X
   Case Is > 0
       Select Case Y
           Case Is > 0
               Atn2 = Atn(Y / X)
           Case Is = 0
               Atn2 = 0
           Case Is < 0
               Atn2 = Atn(Y / X) + Pi + Pi
       End Select
   Case Is = 0
       Select Case Y
           Case Is > 0
               Atn2 = Pi / 2#
           Case Is = 0
               Atn2 = 0
           Case Is < 0
               Atn2 = Pi + Pi / 2#
       End Select
   Case Is < 0
       Atn2 = Pi - Atn(-Y / X)
End Select
Atn2 = Atn2 * 180 / Pi
End Function

Private Sub txtBlue_Change()
txtBlue.Text = Val(txtBlue.Text)
If txtBlue.Text < 0 Then txtBlue.Text = 0
If txtBlue.Text > 255 Then txtBlue.Text = 255
End Sub

Private Sub txtGenScale_Change()
txtGenScale.Text = Val(txtGenScale.Text)
End Sub

Private Sub txtGreen_Change()
txtGreen.Text = Val(txtGreen.Text)
If txtGreen.Text < 0 Then txtGreen.Text = 0
If txtGreen.Text > 255 Then txtGreen.Text = 255
End Sub

Private Sub txtIterations_Change()
txtIterations.Text = Abs(Int(Val(txtIterations.Text)))
End Sub

Private Sub txtRCoord_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then 'enter was pressed
If Val(txtRCoord.Text) > 0 Then
TemplatePoint(ActivePoint).RCoord = txtRCoord.Text
Else
TemplatePoint(ActivePoint).RCoord = 0
End If
TemplatePoint(ActivePoint).XCoord = TemplatePoint(ActivePoint).RCoord * Sin(TemplatePoint(ActivePoint).Theta * Pi / 180)
TemplatePoint(ActivePoint).YCoord = TemplatePoint(ActivePoint).RCoord * -Cos(TemplatePoint(ActivePoint).Theta * Pi / 180)
RedrawScreen
End If
End Sub

Private Sub txtRCoord_LostFocus()
If Val(txtRCoord.Text) > 0 Then
TemplatePoint(ActivePoint).RCoord = txtRCoord.Text
Else
TemplatePoint(ActivePoint).RCoord = 0
End If
TemplatePoint(ActivePoint).XCoord = TemplatePoint(ActivePoint).RCoord * Sin(TemplatePoint(ActivePoint).Theta * Pi / 180)
TemplatePoint(ActivePoint).YCoord = TemplatePoint(ActivePoint).RCoord * -Cos(TemplatePoint(ActivePoint).Theta * Pi / 180)
RedrawScreen
End Sub


Private Sub txtRed_Change()
txtRed.Text = Val(txtRed.Text)
If txtRed.Text < 0 Then txtRed.Text = 0
If txtRed.Text > 255 Then txtRed.Text = 255
End Sub

Private Sub txtRotate_Change()
txtRotate.Text = Val(txtRotate.Text)
End Sub

Private Sub txtScale_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then 'enter was pressed
If Val(txtScale.Text) > 0 Then
TemplatePoint(ActivePoint).ScaleFactor = txtScale.Text
Else
TemplatePoint(ActivePoint).ScaleFactor = 0
End If
RedrawScreen
End If
End Sub

Private Sub txtScale_LostFocus()
If Val(txtScale.Text) > 0 Then
TemplatePoint(ActivePoint).ScaleFactor = txtScale.Text
Else
TemplatePoint(ActivePoint).ScaleFactor = 0
End If
RedrawScreen
End Sub

Private Sub txtTheta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then 'enter was pressed
If Val(txtTheta.Text) > 0 Then
TemplatePoint(ActivePoint).Theta = txtTheta.Text
Else
TemplatePoint(ActivePoint).Theta = 0
End If
With TemplatePoint(ActivePoint)
    .XCoord = TemplatePoint(ActivePoint).RCoord * Sin(TemplatePoint(ActivePoint).Theta * Pi / 180)
    .YCoord = TemplatePoint(ActivePoint).RCoord * -Cos(TemplatePoint(ActivePoint).Theta * Pi / 180)
End With
RedrawScreen
End If
End Sub

Private Sub txtTheta_LostFocus()
If Val(txtTheta.Text) > 0 Then
TemplatePoint(ActivePoint).Theta = txtTheta.Text
Else
TemplatePoint(ActivePoint).Theta = 0
End If
With TemplatePoint(ActivePoint)
    .XCoord = TemplatePoint(ActivePoint).RCoord * Sin(TemplatePoint(ActivePoint).Theta * Pi / 180)
    .YCoord = TemplatePoint(ActivePoint).RCoord * -Cos(TemplatePoint(ActivePoint).Theta * Pi / 180)
End With
RedrawScreen
End Sub

Private Sub txtXCoord_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Val(txtXCoord.Text) > 0 Then
TemplatePoint(ActivePoint).XCoord = txtXCoord.Text
Else
TemplatePoint(ActivePoint).XCoord = 0
End If
With TemplatePoint(ActivePoint)
    .RCoord = Distance((TemplatePoint(ActivePoint).XCoord), (TemplatePoint(ActivePoint).YCoord), 0, 0)
    .Theta = Atn2(-(TemplatePoint(ActivePoint).YCoord), (TemplatePoint(ActivePoint).XCoord))
End With
RedrawScreen
End If
End Sub

Private Sub txtXCoord_LostFocus()
If Val(txtXCoord.Text) > 0 Then
TemplatePoint(ActivePoint).XCoord = txtXCoord.Text
Else
TemplatePoint(ActivePoint).XCoord = 0
End If
With TemplatePoint(ActivePoint)
    .RCoord = Distance((TemplatePoint(ActivePoint).XCoord), (TemplatePoint(ActivePoint).YCoord), 0, 0)
    .Theta = Atn2(-(TemplatePoint(ActivePoint).YCoord), (TemplatePoint(ActivePoint).XCoord))
End With
RedrawScreen
End Sub

Private Sub txtyCoord_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Val(txtYCoord.Text) > 0 Then
TemplatePoint(ActivePoint).YCoord = txtYCoord.Text
Else
TemplatePoint(ActivePoint).YCoord = 0
End If
With TemplatePoint(ActivePoint)
    .RCoord = Distance((TemplatePoint(ActivePoint).XCoord), (TemplatePoint(ActivePoint).YCoord), 0, 0)
    .Theta = Atn2(-(TemplatePoint(ActivePoint).YCoord), (TemplatePoint(ActivePoint).XCoord))
End With
RedrawScreen
End If
End Sub

Private Sub txtyCoord_LostFocus()
If Val(txtYCoord.Text) > 0 Then
TemplatePoint(ActivePoint).YCoord = txtYCoord.Text
Else
TemplatePoint(ActivePoint).YCoord = 0
End If
With TemplatePoint(ActivePoint)
    .RCoord = Distance((TemplatePoint(ActivePoint).XCoord), (TemplatePoint(ActivePoint).YCoord), 0, 0)
    .Theta = Atn2(-(TemplatePoint(ActivePoint).YCoord), (TemplatePoint(ActivePoint).XCoord))
End With
RedrawScreen
End Sub

