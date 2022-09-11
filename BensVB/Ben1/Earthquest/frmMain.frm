VERSION 5.00
Begin VB.Form frmInterface 
   BackColor       =   &H00808080&
   ClientHeight    =   11325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   755
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H0000C0C0&
      Caption         =   "QUIT GAME"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9660
      Width           =   1335
   End
   Begin VB.CommandButton cmdEndTurn 
      BackColor       =   &H000000C0&
      Caption         =   "END TURN"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9060
      Width           =   1335
   End
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00000000&
      Height          =   555
      Left            =   4260
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9600
      Width           =   555
   End
   Begin VB.CommandButton cmdWeapon 
      BackColor       =   &H00404040&
      Height          =   555
      Index           =   3
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9600
      Width           =   555
   End
   Begin VB.CommandButton cmdWeapon 
      BackColor       =   &H00404040&
      Height          =   555
      Index           =   2
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9060
      Width           =   555
   End
   Begin VB.CommandButton cmdWeapon 
      BackColor       =   &H00404040&
      Height          =   555
      Index           =   1
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9600
      Width           =   555
   End
   Begin VB.CommandButton cmdWeapon 
      BackColor       =   &H00404040&
      Height          =   555
      Index           =   0
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9060
      Width           =   555
   End
   Begin VB.Timer tmrTime 
      Interval        =   100
      Left            =   6180
      Top             =   120
   End
   Begin VB.PictureBox picField 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000C000&
      Height          =   8895
      Left            =   60
      MouseIcon       =   "frmMain.frx":0CCA
      MousePointer    =   99  'Custom
      ScaleHeight     =   589
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1001
      TabIndex        =   0
      Top             =   180
      Width           =   15075
      Begin VB.Image imgWeapon 
         Height          =   480
         Index           =   1
         Left            =   360
         Picture         =   "frmMain.frx":1994
         Top             =   3720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgNoDropIcon 
         Height          =   480
         Left            =   -120
         Picture         =   "frmMain.frx":265E
         Top             =   3660
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Label lblMessages 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Game Initialized"
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   8520
      TabIndex        =   10
      Top             =   9060
      Width           =   5235
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   9
      Left            =   0
      Picture         =   "frmMain.frx":3328
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   8
      Left            =   0
      Picture         =   "frmMain.frx":3BF2
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   7
      Left            =   0
      Picture         =   "frmMain.frx":44BC
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   6
      Left            =   480
      Picture         =   "frmMain.frx":4D86
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   5
      Left            =   1140
      Picture         =   "frmMain.frx":5650
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   4
      Left            =   600
      Picture         =   "frmMain.frx":5F1A
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   3
      Left            =   0
      Picture         =   "frmMain.frx":67E4
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   2
      Left            =   60
      Picture         =   "frmMain.frx":74AE
      Top             =   9060
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWeapon 
      Height          =   480
      Index           =   0
      Left            =   1680
      Picture         =   "frmMain.frx":7D78
      Top             =   180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblWeapon 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   855
      Left            =   7380
      TabIndex        =   7
      Top             =   10200
      Width           =   1095
   End
   Begin VB.Image imgUnit 
      Height          =   555
      Left            =   4260
      Top             =   9060
      Width           =   555
   End
   Begin VB.Label lblDisplay 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   1995
      Left            =   4860
      TabIndex        =   2
      Top             =   9060
      Width           =   2475
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTSIZE As Long = &H40
Dim Mode As Integer '1 = move, 2 = attack
Dim ScrollX, ScrollY, BufferDC
'****************************************
Private Sub cmdDraw_Click()

'draw the transparent sprite
BitBlt Me.hdc, 0, 0, SpriteWidth, SpriteHeight, DCMask, 0, 0, vbSrcAnd
BitBlt Me.hdc, 0, 0, SpriteWidth, SpriteHeight, DCSprite, 0, 0, vbSrcPaint

Me.Refresh

End Sub

Private Sub cmdExit_Click()

DeleteGeneratedDC DCMask
DeleteGeneratedDC DCSprite

Unload Me
Set frmMemoryDC = Nothing

End Sub


'IN: FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC < 1 Then
    GenerateDC = 0
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Private Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function

'Private Sub Command1_Click()
'BitBlt Me.hdc, 0, 0, 500, 400, RoomDC(1), 0, 0, vbSrcCopy
'End Sub

Private Sub cmdEndTurn_Click()
Dim i
Redo:
Turn = Turn + 1
If Turn > Teams Then Turn = 1
For i = 1 To Units
Unit(i).Moved = False
Unit(i).Attacked = False
If Unit(i).Owner = Turn And Unit(i).Dead = False Then
ActiveUnit = i
If Unit(i).Attacked = False Then Unit(i).Health = Unit(i).Health + 10
If Unit(i).Health > Unit(i).MaxHP Then Unit(i).Health = Unit(i).MaxHP
End If
Next i
If Unit(ActiveUnit).Owner <> Turn Then GoTo Redo
ActiveWeapon = 0
For i = 0 To 3
cmdWeapon(i).Enabled = True
cmdWeapon(i).Picture = imgWeapon(Unit(ActiveUnit).Weapon(i)).Picture
If Weapon(Unit(ActiveUnit).Weapon(i)).Damage = 0 Or Unit(ActiveUnit).Weapon(i) = 0 Then cmdWeapon(i).Enabled = False
Next i
Mode = 1
ViewX = Unit(ActiveUnit).X - picField.Width / 2
ViewY = Unit(ActiveUnit).Y - picField.Width / 2
If ViewX < 0 Then ViewX = 0
If ViewX > 1000 Then ViewX = 1000
If ViewY < 0 Then ViewY = 0
If ViewY > 1000 Then ViewY = 1000
Call RedrawScreen
Call RefreshText
End Sub

Private Sub cmdMove_Click()
ActiveWeapon = 0
Mode = 1 'move
End Sub

Private Sub cmdQuit_Click()
Unload Me
End
End Sub

Private Sub cmdWeapon_Click(Index As Integer)
ActiveSlot = Index
ActiveWeapon = Unit(ActiveUnit).Weapon(Index)
Mode = 2 'attack
End Sub

Private Sub Form_Load()
Dim i, X
Call InitiateUnits 'load unit data into memory
Turn = 1
'Call LoadUnit(1, 400, 100, 1, 1, 0, 0, 0)
'Call LoadUnit(1, 100, 100, 1, 2, 1, 0, 0)
'Call LoadUnit(1, 100, 400, 2, 2, 1, 0, 0)
'Call LoadUnit(1, 400, 400, 2, 1, 0, 0, 0)
'Call LoadUnit(1, 700, 100, 1, 3, 3, 0, 0)
'Call LoadUnit(1, 700, 400, 2, 3, 2, 0, 0)
For i = 1 To 1
For X = 1 To 8
UnitDC(i, X) = GenerateDC(App.Path & "\Unit" & i & X & ".bmp")
MaskDC(i, X) = GenerateDC(App.Path & "\Mask" & i & X & ".bmp")
Next X
Next i
'BufferDC = GenerateDC(App.Path & ")
Turn = 1
Teams = 2
ActiveUnit = 1
InitiateUnits 'load unit data into memory
Randomize
For i = 0 To 3
cmdWeapon(i).Enabled = True
cmdWeapon(i).Picture = imgWeapon(Unit(ActiveUnit).Weapon(i)).Picture
If Weapon(Unit(ActiveUnit).Weapon(i)).Damage = 0 Or Unit(ActiveUnit).Weapon(i) = 0 Then cmdWeapon(i).Enabled = False
Next i
Mode = 1
Call RedrawScreen
Call RefreshText
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case X
Case Is < 3
ScrollX = -10
Case Is > 1013
ScrollX = 6
Case Else
ScrollX = 0
End Select
Select Case Y
Case Is < 3
ScrollY = -10
Case Is > 738
ScrollY = 6
Case Else
ScrollY = 0
End Select
End Sub

Private Sub picField_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Mode
Case Is = 1
If CheckRange(X, Y, Unit(ActiveUnit).X - ViewX, Unit(ActiveUnit).Y - ViewY, (Unit(ActiveUnit).Movement)) = True And Unit(ActiveUnit).Moved = False Then
picField.MouseIcon = cmdMove.Picture
Else
picField.MouseIcon = imgNoDropIcon.Picture
End If
Case Is = 2
If CheckRange(X, Y, Unit(ActiveUnit).X - ViewX, Unit(ActiveUnit).Y - ViewY, (Weapon(ActiveWeapon).MaxRange)) = True And Unit(ActiveUnit).Attacked = False Then
picField.MouseIcon = cmdWeapon(ActiveSlot).Picture
Else
picField.MouseIcon = imgNoDropIcon.Picture
End If
End Select
ScrollX = 0
ScrollY = 0
End Sub

Private Sub picField_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i, n
'<<<CHECK WHETHER FRIENDLY UNIT IS BEING CLICKED>>>
If Button = 1 Then 'left click
For i = 1 To Units
    If Unit(i).Owner = Turn And Unit(i).Dead = False Then 'if the unit belongs to the player whose turn it is
        If CheckRange(X, Y, Unit(i).X - ViewX, Unit(i).Y - ViewY, (Unit(i).Size)) = True Then
        ActiveUnit = i
        Mode = 1
        ActiveWeapon = 0
        For n = 0 To 3
        cmdWeapon(n).Enabled = True
        cmdWeapon(n).Picture = imgWeapon(Unit(ActiveUnit).Weapon(n)).Picture
        If Weapon(Unit(ActiveUnit).Weapon(n)).Damage = 0 Or Unit(ActiveUnit).Weapon(n) = 0 Then cmdWeapon(n).Enabled = False
        Next n
        End If
    End If
Next i
ViewX = X + ViewX - picField.Width / 2
ViewY = Y + ViewY - picField.Height / 2
If ViewX < 0 Then ViewX = 0
If ViewX > 1000 Then ViewX = 1000
If ViewY < 0 Then ViewY = 0
If ViewY > 1000 Then ViewY = 1000
Else 'right click -- move active unit or attack a clicked one with selected weapon
If ActiveWeapon = 0 And Unit(ActiveUnit).Moved = False Then 'if the move tool is selected
    For i = 1 To Units
        If CheckRange(X, Y, Unit(i).X - ViewX, Unit(i).Y - ViewY, Unit(i).Size + Unit(ActiveUnit).Size) = True And i <> ActiveUnit Then Exit Sub
        'can't move unit onto already occupied space
    Next i
    If CheckRange(X, Y, Unit(ActiveUnit).X - ViewX, Unit(ActiveUnit).Y - ViewY, (Unit(ActiveUnit).Movement)) = True Then
        Unit(ActiveUnit).X = X + ViewX
        Unit(ActiveUnit).Y = Y + ViewY
        Unit(ActiveUnit).Moved = True
    End If
End If
If ActiveWeapon > 0 And Unit(ActiveUnit).Attacked = False Then 'if a weapon is selected
    For i = 1 To Units
        If Unit(i).Owner <> Turn And Unit(i).Dead = False Then 'if the unit is an enemy
        If CheckRange(X, Y, Unit(i).X - ViewX, Unit(i).Y - ViewY, (Unit(i).Size)) = True Then
            If CheckRange(X, Y, Unit(ActiveUnit).X - ViewX, Unit(ActiveUnit).Y - ViewY, (Weapon(ActiveWeapon).MaxRange)) = True Then
                Call Attack(ActiveUnit, i)
                Unit(ActiveUnit).Attacked = True
                Exit For
            End If
        End If
        End If
    Next i
End If
End If
Call RefreshText
End Sub



Public Sub RedrawScreen()
Dim i, PI
PI = Atn(1) * 4
picField.Cls
For i = 1 To Units
Select Case Unit(i).Owner
Case Is = 1
If Unit(i).Dead = False Then picField.Circle (Unit(i).X - ViewX, Unit(i).Y - ViewY), Unit(i).Size, RGB(240, 0, 0)
Case Is = 2
If Unit(i).Dead = False Then picField.Circle (Unit(i).X - ViewX, Unit(i).Y - ViewY), Unit(i).Size, RGB(0, 10, 240)
End Select
If Unit(i).Dead = False Then picField.Circle (Unit(i).X - ViewX, Unit(i).Y - ViewY), Unit(i).Size + 2, RGB(0, 255, 0), 225 * PI / 180, (225 + 90 * Unit(i).Health / Unit(i).MaxHP) * PI / 180
Next i
picField.Circle (Unit(ActiveUnit).X - ViewX, Unit(ActiveUnit).Y - ViewY), Unit(ActiveUnit).Size - 2, RGB(255, 255, 50)
For i = 1 To Units
BitBlt picField.hdc, Unit(i).X - Unit(i).Size - ViewX, Unit(i).Y - Unit(i).Size - ViewY, Unit(i).Size * 2, Unit(i).Size * 2, MaskDC(Unit(i).GraphicsDC, Unit(i).Position), 0, 0, vbSrcAnd
BitBlt picField.hdc, Unit(i).X - Unit(i).Size - ViewX, Unit(i).Y - Unit(i).Size - ViewY, Unit(i).Size * 2, Unit(i).Size * 2, UnitDC(Unit(i).GraphicsDC, Unit(i).Position), 0, 0, vbSrcPaint
Next i
End Sub

Private Sub tmrTime_Timer()
RedrawScreen
ViewX = ViewX + ScrollX
ViewY = ViewY + ScrollY
If ViewX < 0 Then ViewX = 0
If ViewX > 1000 Then ViewX = 1000
If ViewY < 0 Then ViewY = 0
If ViewY > 1000 Then ViewY = 1000
End Sub

Public Sub RefreshText()
Dim i
lblDisplay.Caption = Unit(ActiveUnit).Name & vbLf & _
    "Level " & Unit(ActiveUnit).level & vbLf & _
    "HP: " & Unit(ActiveUnit).Health & "/" & Unit(ActiveUnit).MaxHP
For i = 0 To 3
If Weapon(Unit(ActiveUnit).Weapon(i)).Damage > 0 Then
lblDisplay.Caption = lblDisplay.Caption & vbLf & Weapon(Unit(ActiveUnit).Weapon(i)).Name & _
    "   XP: " & Int(Unit(ActiveUnit).Skill(i) * 100)
End If
Next i
End Sub
