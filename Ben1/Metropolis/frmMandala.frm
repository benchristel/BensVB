VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   11475
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   15330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   765
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1022
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H000000FF&
      Caption         =   "Quit"
      Height          =   255
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdRollDice 
      Caption         =   "Roll Dice"
      Height          =   375
      Left            =   60
      TabIndex        =   41
      Top             =   10440
      Width           =   915
   End
   Begin VB.TextBox txtDicePlus 
      Height          =   315
      Left            =   1080
      TabIndex        =   36
      Text            =   "1"
      Top             =   10020
      Width           =   915
   End
   Begin VB.TextBox txtDiceSides 
      Height          =   315
      Left            =   1080
      TabIndex        =   35
      Text            =   "6"
      Top             =   9660
      Width           =   915
   End
   Begin VB.TextBox txtDice 
      Height          =   315
      Left            =   1080
      TabIndex        =   34
      Text            =   "1"
      Top             =   9300
      Width           =   915
   End
   Begin VB.TextBox txtMoveRestrict 
      Height          =   315
      Left            =   2040
      TabIndex        =   32
      Top             =   9960
      Width           =   1335
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Left            =   14520
      TabIndex        =   29
      Top             =   9300
      Width           =   675
   End
   Begin VB.CommandButton cmdPlaceUnit 
      Caption         =   "Place Unit"
      Height          =   315
      Left            =   13740
      TabIndex        =   28
      Top             =   10380
      Width           =   1515
   End
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   2940
      Top             =   9300
   End
   Begin VB.PictureBox picUnitGraphics 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2115
      Left            =   3420
      ScaleHeight     =   137
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   27
      Top             =   9300
      Width           =   2115
   End
   Begin VB.CommandButton cmdDefaultLabels 
      Caption         =   "Set Default Labels"
      Height          =   315
      Left            =   13740
      TabIndex        =   26
      Top             =   10740
      Width           =   1515
   End
   Begin VB.CommandButton cmdLockLabels 
      Caption         =   "Lock Labels"
      Height          =   315
      Left            =   13740
      TabIndex        =   25
      Top             =   11100
      Width           =   1515
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   11
      Left            =   9660
      TabIndex        =   24
      Top             =   11100
      Width           =   1335
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   10
      Left            =   9660
      TabIndex        =   23
      Top             =   10740
      Width           =   1335
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   9
      Left            =   9660
      TabIndex        =   22
      Top             =   10380
      Width           =   1335
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   8
      Left            =   9660
      TabIndex        =   21
      Top             =   10020
      Width           =   1335
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   7
      Left            =   9660
      TabIndex        =   20
      Top             =   9660
      Width           =   1335
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   6
      Left            =   9660
      TabIndex        =   19
      Top             =   9300
      Width           =   1335
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   5
      Left            =   5580
      TabIndex        =   18
      Top             =   11100
      Width           =   1335
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   4
      Left            =   5580
      TabIndex        =   17
      Top             =   10740
      Width           =   1335
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   3
      Left            =   5580
      TabIndex        =   16
      Top             =   10380
      Width           =   1335
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   2
      Left            =   5580
      TabIndex        =   15
      Top             =   10020
      Width           =   1335
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   1
      Left            =   5580
      TabIndex        =   14
      Top             =   9660
      Width           =   1335
   End
   Begin VB.TextBox txtLabel 
      Height          =   315
      Index           =   0
      Left            =   5580
      TabIndex        =   13
      Top             =   9300
      Width           =   1335
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   11
      Left            =   11040
      TabIndex        =   12
      Top             =   11100
      Width           =   2655
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   10
      Left            =   11040
      TabIndex        =   11
      Top             =   10740
      Width           =   2655
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   9
      Left            =   11040
      TabIndex        =   10
      Top             =   10380
      Width           =   2655
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   8
      Left            =   11040
      TabIndex        =   9
      Top             =   10020
      Width           =   2655
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   7
      Left            =   11040
      TabIndex        =   8
      Top             =   9660
      Width           =   2655
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   6
      Left            =   11040
      TabIndex        =   7
      Top             =   9300
      Width           =   2655
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   5
      Left            =   6960
      TabIndex        =   6
      Top             =   11100
      Width           =   2655
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   4
      Left            =   6960
      TabIndex        =   5
      Top             =   10740
      Width           =   2655
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   3
      Left            =   6960
      TabIndex        =   4
      Top             =   10380
      Width           =   2655
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   2
      Left            =   6960
      TabIndex        =   3
      Top             =   10020
      Width           =   2655
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   1
      Left            =   6960
      TabIndex        =   2
      Top             =   9660
      Width           =   2655
   End
   Begin VB.TextBox txtProperty 
      Height          =   315
      Index           =   0
      Left            =   6960
      TabIndex        =   1
      Top             =   9300
      Width           =   2655
   End
   Begin VB.PictureBox picField 
      Height          =   8775
      Left            =   120
      MouseIcon       =   "frmMandala.frx":0000
      ScaleHeight     =   581
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1009
      TabIndex        =   0
      Top             =   360
      Width           =   15195
   End
   Begin VB.Label lblDiceRoll 
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   40
      Top             =   10440
      Width           =   915
   End
   Begin VB.Label lblDice 
      Alignment       =   1  'Right Justify
      Caption         =   "Add to Roll"
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   39
      Top             =   10080
      Width           =   975
   End
   Begin VB.Label lblDice 
      Alignment       =   1  'Right Justify
      Caption         =   "Dice Sides"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   38
      Top             =   9720
      Width           =   975
   End
   Begin VB.Label lblDice 
      Alignment       =   1  'Right Justify
      Caption         =   "# Of Dice"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   37
      Top             =   9360
      Width           =   975
   End
   Begin VB.Label lblRestrictMoves 
      Caption         =   "Restrict Moves By:"
      Height          =   255
      Left            =   2040
      TabIndex        =   33
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label lblDistance 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      TabIndex        =   31
      Top             =   9300
      Width           =   1275
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   13740
      TabIndex        =   30
      Top             =   9300
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Unit() As Unit, ActiveUnit As Integer, Units As Integer
Dim DC(1 To 10, 1 To 4) As Long, PlaceDC As Integer, GraphicsSize(1 To 10) As Integer
Dim MaskDC(1 To 10, 1 To 4) As Long
Dim CreateUnit As Boolean
Dim BackBuffDC As Long, BackGroundDC As Long, BaseDC(0 To 1) As Long, BaseMaskDC As Long
Dim ViewX As Single, ViewY As Single
Dim PrintDistance As Integer, MaxMovement As Integer
Dim MoveViewX As Integer, MoveViewY As Integer
Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub cmdDefaultLabels_Click()
txtLabel(0).Text = "Name"
txtLabel(1).Text = "Hitpoints"
txtLabel(2).Text = "Attack Dice"
txtLabel(3).Text = "Attack Dice Sides"
txtLabel(4).Text = "Attack +"
txtLabel(5).Text = "Save Dice"
txtLabel(6).Text = "Save Dice Sides"
txtLabel(7).Text = "Save +"
txtLabel(8).Text = "Movement"
txtLabel(9).Text = "Range"
txtMoveRestrict.Text = 9
End Sub

Private Sub cmdLockLabels_Click()
Dim i
Select Case cmdLockLabels.Caption
Case Is = "Lock Labels"
For i = 0 To 11
txtLabel(i).Enabled = False
Next i
cmdDefaultLabels.Enabled = False
cmdLockLabels.Caption = "Unlock Labels"
Case Is = "Unlock Labels"
For i = 0 To 11
txtLabel(i).Enabled = True
Next i
cmdDefaultLabels.Enabled = True
cmdLockLabels.Caption = "Lock Labels"
End Select
End Sub

Private Sub cmdPlaceUnit_Click()
If CreateUnit = False Then
CreateUnit = True
cmdPlaceUnit.Caption = "Normal Play"
Else
CreateUnit = False
cmdPlaceUnit.Caption = "Place Unit"
End If
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub cmdRollDice_Click()
lblDiceRoll.Caption = RollDice(Int(Val(txtDice.Text)), Int(Val(txtDiceSides.Text)), Int(Val(txtDicePlus.Text)))
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim i, j
Randomize
For i = 1 To 10
For j = 1 To 4
DC(i, j) = GenerateDC(App.Path & "\unit" & i & j & ".bmp")
MaskDC(i, j) = GenerateDC(App.Path & "\mask" & i & j & ".bmp")
Next j
Next i
For i = 0 To 1
BaseDC(i) = GenerateDC(App.Path & "\base" & i & ".bmp")
Next i
BaseMaskDC = GenerateDC(App.Path & "\basemask.bmp")
BackBuffDC = GenerateDC(App.Path & "\Background.bmp")
BackGroundDC = GenerateDC(App.Path & "\Background.bmp")
PlaceDC = 1
GraphicsSize(1) = 48
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case x
    Case Is < 3
    MoveViewX = 20
    Case Is > Me.ScaleWidth - 3
    MoveViewX = -20
    Case Else
    MoveViewX = 0
End Select
Select Case y
    Case Is < 3
    MoveViewY = 20
    Case Is > Me.ScaleHeight - 3
    MoveViewY = -20
    Case Else
    MoveViewY = 0
End Select
End Sub

Private Sub Form_Terminate()
Dim i, j
For i = 1 To 10
For j = 1 To 4
DeleteGeneratedDC (DC(i, j))
DeleteGeneratedDC (MaskDC(i, j))
Next j
Next i
For i = 0 To 1
DeleteGeneratedDC (BaseDC(i))
Next i
DeleteGeneratedDC (BackBuffDC)
DeleteGeneratedDC (BackGroundDC)
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub picField_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i, j
If Button = 1 Then
    ActiveUnit = 0
    For i = 1 To Units
    If distance(Unit(i).x, Unit(i).y, x - ViewX, y - ViewY) <= Unit(i).Size And CreateUnit = False Then
    ActiveUnit = i
    If Int(Val(txtMoveRestrict.Text)) > 0 And Int(Val(txtMoveRestrict.Text)) < 13 Then
    MaxMovement = txtProperty(Int(Val(txtMoveRestrict.Text)) - 1).Text
    Else
    MaxMovement = 0
    End If
    For j = 0 To 11
    txtProperty(j).Text = Unit(i).Stats(j)
    Next j
    End If
    Next i
    If CreateUnit = True Then
    Units = Units + 1
    ReDim Preserve Unit(1 To Units)
    With Unit(Units)
        .GraphicsDC = PlaceDC
    For i = 0 To 11
        .Stats(i) = txtProperty(i).Text
    Next i
        .x = x - ViewX
        .y = y - ViewY
        .Size = GraphicsSize(PlaceDC)
        .Position = Int(Rnd * 4 + 1)
        .BaseType = 1
        .Size = Val(txtSize.Text)
    End With
        If Val(txtSize.Text) = 0 Then Unit(Units).Size = GraphicsSize(Unit(Units).GraphicsDC) / 2
        ActiveUnit = Units
    End If
Else
    If CreateUnit = False And ActiveUnit > 0 Then
    Unit(ActiveUnit).x = x - ViewX
    Unit(ActiveUnit).y = y - ViewY
    End If
End If
End Sub

Private Sub picField_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim cursordistance As Integer
    If ActiveUnit > 0 And CreateUnit = False Then
    cursordistance = Int(distance(Unit(ActiveUnit).x, Unit(ActiveUnit).y, x - ViewX, y - ViewY))
    End If
    lblDistance.Caption = cursordistance
    If MaxMovement > 0 And CreateUnit = False Then
    If cursordistance > MaxMovement Then
    picField.MousePointer = 99
    Else
    picField.MousePointer = 0
    End If
    End If
MoveViewX = 0
MoveViewY = 0
    
End Sub

Private Sub tmrTime_Timer()
Dim i
'move view
ViewX = ViewX + MoveViewX
ViewY = ViewY + MoveViewY
If ViewX > 0 Then ViewX = 0
If ViewX < -180 Then ViewX = -180
If ViewY > 0 Then ViewY = 0
If ViewY < -620 Then ViewY = -620
    'blit background to buffer
BitBlt BackBuffDC, 0, 0, picField.ScaleWidth, picField.ScaleHeight, BackGroundDC, -ViewX, -ViewY, vbSrcCopy
    'Blit Graphics to backbuffer
For i = 1 To Units
    'blit base graphics
    StretchBlt BackBuffDC, Unit(i).x + ViewX - Unit(i).Size, Unit(i).y + ViewY - Unit(i).Size, Unit(i).Size * 2, Unit(i).Size * 2, BaseMaskDC, 0, 0, 140, 140, vbSrcAnd
If ActiveUnit = i Then
    StretchBlt BackBuffDC, Unit(i).x + ViewX - Unit(i).Size, Unit(i).y + ViewY - Unit(i).Size, Unit(i).Size * 2, Unit(i).Size * 2, BaseDC(0), 0, 0, 140, 140, vbSrcPaint
Else
    StretchBlt BackBuffDC, Unit(i).x + ViewX - Unit(i).Size, Unit(i).y + ViewY - Unit(i).Size, Unit(i).Size * 2, Unit(i).Size * 2, BaseDC(Unit(i).BaseType), 0, 0, 140, 140, vbSrcPaint
End If
    'blit unit graphics
    BitBlt BackBuffDC, Unit(i).x + ViewX - GraphicsSize(Unit(i).GraphicsDC) / 2, Unit(i).y + ViewY - GraphicsSize(Unit(i).GraphicsDC) / 2, GraphicsSize(Unit(i).GraphicsDC), GraphicsSize(Unit(i).GraphicsDC), MaskDC(Unit(i).GraphicsDC, Unit(i).Position), 0, 0, vbSrcAnd
    BitBlt BackBuffDC, Unit(i).x + ViewX - GraphicsSize(Unit(i).GraphicsDC) / 2, Unit(i).y + ViewY - GraphicsSize(Unit(i).GraphicsDC) / 2, GraphicsSize(Unit(i).GraphicsDC), GraphicsSize(Unit(i).GraphicsDC), DC(Unit(i).GraphicsDC, Unit(i).Position), 0, 0, vbSrcPaint
Next i
    'blit backbuffer to screen
    BitBlt picField.hdc, 0, 0, picField.ScaleWidth, picField.ScaleHeight, BackBuffDC, 0, 0, vbSrcCopy
    picUnitGraphics.Cls
    BitBlt picUnitGraphics.hdc, 68.5 - GraphicsSize(PlaceDC) / 2, 68.5 - GraphicsSize(PlaceDC) / 2, GraphicsSize(PlaceDC), GraphicsSize(PlaceDC), DC(PlaceDC, 1), 0, 0, vbSrcCopy
    picUnitGraphics.Refresh
End Sub

Private Sub txtMoveRestrict_Change()
    If Int(Val(txtMoveRestrict.Text)) > 0 And Int(Val(txtMoveRestrict.Text)) < 13 Then
    MaxMovement = Int(Val(txtProperty(Int(Val(txtMoveRestrict.Text)) - 1).Text))
    Else
    MaxMovement = 0
    End If
End Sub

Private Sub txtProperty_Change(Index As Integer)
If ActiveUnit = 0 Then Exit Sub
Unit(ActiveUnit).Stats(Index) = txtProperty(Index).Text
If Int(Val(txtMoveRestrict.Text)) = Index Then MaxMovement = Int(Val(txtProperty(Index).Text))
End Sub

Private Sub txtProperty_KeyPress(Index As Integer, KeyAscii As Integer)
'THIS SUB WAS SUPPOSED TO MAKE THE COMPUTER EVALUATE AN EXPRESSION TYPED INTO A
'PROPERTY TEXTBOX TO MAKE IT AUTOMATICALLY PERFORM TYPED CALCULATIONS BUT IT DIDN'T WORK.
'INSERT ALTERNATE CODE LATER???
'If KeyAscii = 13 Then 'enter
'txtProperty(Index).Text = Int(txtProperty(Index).Text)
'End If
End Sub

Private Sub txtSize_Change()
If ActiveUnit > 0 Then
    Unit(ActiveUnit).Size = Val(txtSize.Text)
    If Val(txtSize.Text) = 0 Then Unit(ActiveUnit).Size = GraphicsSize(Unit(ActiveUnit).GraphicsDC)
End If
End Sub
