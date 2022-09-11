VERSION 5.00
Begin VB.Form frmBattleground 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   5730
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   15270
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   5730
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picDead 
      Height          =   555
      Index           =   5
      Left            =   8940
      Picture         =   "frmBattleground.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   660
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picDead 
      Height          =   555
      Index           =   4
      Left            =   8340
      Picture         =   "frmBattleground.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   660
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picDead 
      Height          =   555
      Index           =   3
      Left            =   7740
      Picture         =   "frmBattleground.frx":1194
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   660
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picDead 
      Height          =   555
      Index           =   2
      Left            =   7140
      Picture         =   "frmBattleground.frx":1A5E
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   660
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picDead 
      Height          =   555
      Index           =   1
      Left            =   6540
      Picture         =   "frmBattleground.frx":2328
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      Top             =   660
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picDead 
      Height          =   555
      Index           =   0
      Left            =   5940
      Picture         =   "frmBattleground.frx":2BF2
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   660
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4260
      Top             =   60
   End
   Begin VB.PictureBox picUnit 
      BackColor       =   &H00000000&
      Height          =   555
      Index           =   5
      Left            =   8940
      Picture         =   "frmBattleground.frx":34BC
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picUnit 
      BackColor       =   &H00000000&
      Height          =   555
      Index           =   4
      Left            =   8340
      Picture         =   "frmBattleground.frx":3D86
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picUnit 
      BackColor       =   &H00000000&
      Height          =   555
      Index           =   3
      Left            =   7740
      Picture         =   "frmBattleground.frx":4650
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picUnit 
      BackColor       =   &H00000000&
      Height          =   555
      Index           =   2
      Left            =   7140
      Picture         =   "frmBattleground.frx":4F1A
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picUnit 
      BackColor       =   &H00000000&
      Height          =   555
      Index           =   1
      Left            =   6540
      Picture         =   "frmBattleground.frx":57E4
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picUnit 
      BackColor       =   &H00000000&
      Height          =   555
      Index           =   0
      Left            =   5940
      Picture         =   "frmBattleground.frx":60AE
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox picMask 
      BackColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   5340
      Picture         =   "frmBattleground.frx":6978
      ScaleHeight     =   33
      ScaleMode       =   0  'User
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblCasualties 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   13380
      TabIndex        =   23
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblCasualties 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   13380
      TabIndex        =   22
      Top             =   1320
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblCasualties 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   13380
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblCasualties 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   13380
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblCasualties 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   13380
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblCasualties 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   13380
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblCasualties 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   13380
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label lblTrackAttack 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   60
      TabIndex        =   16
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label lblTrackMode 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   60
      TabIndex        =   15
      Top             =   420
      Width           =   1755
   End
   Begin VB.Label lblTrackHP 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   240
      Width           =   1755
   End
   Begin VB.Label lblTrackType 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   60
      TabIndex        =   13
      Top             =   60
      Width           =   1755
   End
   Begin VB.Image imgUnit 
      Height          =   555
      Index           =   0
      Left            =   4740
      Top             =   60
      Width           =   555
   End
   Begin VB.Shape shpTrackHilite 
      BorderColor     =   &H000080FF&
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   5340
      Shape           =   3  'Circle
      Top             =   660
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Battle"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Battle"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuBattle 
      Caption         =   "Battle"
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Units..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSelectTeam 
         Caption         =   "Select Team"
         Begin VB.Menu mnuTeam 
            Caption         =   "&Death"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuTeam 
            Caption         =   "&Life"
            Index           =   1
         End
         Begin VB.Menu mnuTeam 
            Caption         =   "&Calm"
            Index           =   2
         End
         Begin VB.Menu mnuTeam 
            Caption         =   "&Blood"
            Index           =   3
         End
         Begin VB.Menu mnuTeam 
            Caption         =   "&Plague"
            Index           =   4
         End
         Begin VB.Menu mnuTeam 
            Caption         =   "&Might"
            Index           =   5
         End
      End
      Begin VB.Menu mnuStart 
         Caption         =   "Start Battle"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "End Battle"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuDisplay 
      Caption         =   "Display"
      Begin VB.Menu mnuSelectColor 
         Caption         =   "Select Color"
         Begin VB.Menu mnuColor1 
            Caption         =   "Misty Cliffs"
         End
         Begin VB.Menu mnuColor2 
            Caption         =   "Bloodstained"
         End
         Begin VB.Menu mnuColor3 
            Caption         =   "Sunny"
         End
         Begin VB.Menu mnuColor4 
            Caption         =   "Forest"
         End
         Begin VB.Menu mnuColor5 
            Caption         =   "Poseidon"
         End
         Begin VB.Menu mnuColor6 
            Caption         =   "Deep Sea"
         End
      End
      Begin VB.Menu mnuCasualties 
         Caption         =   "Casualties"
      End
      Begin VB.Menu mnuTracking 
         Caption         =   "Tracking"
         Begin VB.Menu mnuTrackEnabled 
            Caption         =   "Enabled"
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuSelectTrack 
            Caption         =   "Select"
            Shortcut        =   {F4}
         End
      End
   End
End
Attribute VB_Name = "frmBattleground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Firepower(1 To 100, 6) As Integer, Range(1 To 100, 6) As Integer, Armor(1 To 100, 6) As Integer
Dim Health(1 To 100) As Integer, HP(1 To 100) As Integer, Chicken(1 To 100) As Integer, _
    Speed(1 To 100) As Integer, Threshold(1 To 100) As Integer, Healing(1 To 100) As Integer, _
    RateOfFire(1 To 100) As Integer
Dim Team(1 To 100) As Integer
Dim Units As Integer, UnitName(1 To 100) As String
Dim UnitMode(1 To 100) As Integer, UnitX(1 To 100) As Integer, UnitY(1 To 100) As Integer
Dim UnitPos(1 To 100) As Integer
Dim PColor As Integer
Dim Casualties(0 To 6) As Integer
Dim Tracking As Boolean, TrackNo As Integer
Private Declare Function BitBlt Lib "gdi32" _
                (ByVal hDestDC As Long, _
                 ByVal X As Long, _
                 ByVal Y As Long, _
                 ByVal nWidth As Long, _
                 ByVal nHeight As Long, _
                 ByVal hSrcDC As Long, _
                 ByVal xSrc As Long, _
                 ByVal ySrc As Long, _
                 ByVal dwRop As Long) As Long

Private Sub Form_Load()
' I dunno...  GenerateDC ("C:/MyDocuments/Ben/BensVB/Ben1/Metropolis/TrooperBlue.ico")
PColor = 0 'black
'1 = white
'2 = blue
'3 = red
'4 = yellow
'5 = purple
PFirepower(0) = 2
PFirepower(1) = 3
PFirepower(2) = 1
PFirepower(3) = 0
PFirepower(4) = 3
PFirepower(5) = 3
PFirepower(6) = 0
PRange(0) = 480
PRange(1) = 600
PRange(2) = 480
PRange(3) = 0
PRange(4) = 450
PRange(5) = 450
PRange(6) = 0
PArmor(0) = 10
PArmor(1) = 10
PArmor(2) = 10
PArmor(3) = 10
PArmor(4) = 5
PArmor(5) = 5
PArmor(6) = 0
PHealth = 600
PHP = 600
PChicken = 360
PSpeed = 35
PThreshold = 500
PHealing = 2
PRateOfFire = 2
TrackNo = 1
PName = "Militia"
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i
If Units = 100 Then
MsgBox "You can't add any more units.", , "Unit Limit"
Exit Sub
End If
'BitBlt Me.hDC, X - 16, Y - 16, picMask.ScaleWidth, picMask.ScaleHeight, picMask.hDC, 0, 0, vbSrcAnd
'BitBlt Me.hDC, X - 16, Y - 16, picUnit(PColor).ScaleWidth, picUnit(PColor).ScaleHeight, picUnit(PColor).hDC, 0, 0, vbSrcPaint
Units = Units + 1
For i = 0 To 6
Firepower(Units, i) = PFirepower(i)
Armor(Units, i) = PArmor(i)
Range(Units, i) = PRange(i)
Next i
Health(Units) = PHealth
HP(Units) = PHP
Chicken(Units) = PChicken
Speed(Units) = PSpeed
Threshold(Units) = PThreshold
Healing(Units) = PHealing
RateOfFire(Units) = PRateOfFire
UnitX(Units) = X - 240
UnitY(Units) = Y - 240
UnitMode(Units) = 2 'attack
Team(Units) = PColor
UnitName(Units) = PName
Load imgUnit(Units)
imgUnit(Units).Picture = picUnit(PColor).Picture
imgUnit(Units).Left = UnitX(Units)
imgUnit(Units).Top = UnitY(Units)
imgUnit(Units).Visible = True
imgUnit(Units).ZOrder
lblTrackType.ZOrder
lblTrackHP.ZOrder
lblTrackMode.ZOrder
lblTrackAttack.ZOrder
If Tracking = True Then
shpTrackHilite.Visible = True
shpTrackHilite.Move UnitX(TrackNo), UnitY(TrackNo)
End If
End Sub

Private Sub mnuCasualties_Click()
Dim i
If mnuCasualties.Checked = False Then
mnuCasualties.Checked = True
For i = 0 To 6
lblCasualties(i).Visible = True
Next i
Else
mnuCasualties.Checked = False
For i = 0 To 6
lblCasualties(i).Visible = False
Next i
End If
End Sub

Private Sub mnuEdit_Click()
Load frmEditUnits
frmEditUnits.Visible = True
frmEditUnits.ZOrder
End Sub

Private Sub mnuEnd_Click()
Dim i, X
tmrTime.Enabled = False
For i = 1 To Units
Unload imgUnit(i)
For X = 0 To 6
Firepower(i, X) = 0
Armor(i, X) = 0
Range(i, X) = 0
Next X
Health(i) = 0
HP(i) = 0
Chicken(i) = 0
Speed(i) = 0
Threshold(i) = 0
Healing(i) = 0
UnitX(i) = 0
UnitY(i) = 0
UnitMode(i) = 0
Team(i) = 0
Next i
Units = 0
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub
' Public Function GenerateDC(FileName As String) As Long
'Dim DC As Long
'Dim hBitmap As Long
'
''Create a Device Context, compatible with the screen
'DC = CreateCompatibleDC(0)
'
'If DC < 1 Then
'    GenerateDC = 0
'    Exit Function
'End If
'
''Load the image....
''BIG NOTE: This function is not supported under NT, there you can not
''specify the LR_LOADFROMFILE flag
'
'hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
'
'If hBitmap = 0 Then 'Failure in loading bitmap
'    DeleteDC DC
'    GenerateDC = 0
'    Exit Function
'End If
'
''Assign the Bitmap to the Device Context
'SelectObject DC, hBitmap
'
''Return the device context
'GenerateDC = DC
'
''Delete the bitmap handle object
'DeleteObject hBitmap
'
'End Function
'

Private Sub mnuPause_Click()
tmrTime.Enabled = False
MsgBox "CLICK OK TO CONTINUE", , "PAUSED"
tmrTime.Enabled = True
End Sub

Private Sub mnuSelectTrack_Click()
redo:
If TrackNo < Units Then
TrackNo = TrackNo + 1
Else
TrackNo = 1
End If
If UnitMode(TrackNo) = 0 Then GoTo redo
shpTrackHilite.Move UnitX(TrackNo), UnitY(TrackNo)
End Sub

Private Sub mnuStart_Click()
If Units > 0 Then
tmrTime.Enabled = True
Else
MsgBox "You haven't placed any units yet.", , "Error"
End If
End Sub

Private Sub mnuTeam_Click(Index As Integer)
Dim i
PColor = Index
For i = 0 To 5
mnuTeam(i).Checked = False
Next i
mnuTeam(Index).Checked = True
End Sub

Private Sub Picture1_Click()

End Sub


Private Sub mnuTrackEnabled_Click()
If mnuTrackEnabled.Checked = True Then
mnuTrackEnabled.Checked = False
Tracking = False
lblTrackType.Visible = False
lblTrackHP.Visible = False
lblTrackMode.Visible = False
lblTrackAttack.Visible = False
shpTrackHilite.Visible = False
Else
mnuTrackEnabled.Checked = True
Tracking = True
lblTrackType.Visible = True
lblTrackHP.Visible = True
lblTrackMode.Visible = True
lblTrackAttack.Visible = True
If Units > 0 Then
shpTrackHilite.Visible = True
shpTrackHilite.Move UnitX(TrackNo), UnitY(TrackNo)
End If
End If
End Sub

Private Sub tmrTime_Timer()
Dim i, X, a, XAdd, YAdd, Attacked(1 To 100) As Integer, UnitAttacked As Boolean
For i = 1 To Units
'Attacked = 0
If HP(i) < Chicken(i) And UnitMode(i) <> 0 Then UnitMode(i) = 1
If HP(i) > Threshold(i) And UnitMode(i) <> 0 Then UnitMode(i) = 2
If HP(i) < 0 Then
HP(i) = 0
UnitMode(i) = 0 'dead
imgUnit(i).Picture = picDead(Team(i)).Picture
Casualties(Team(i)) = Casualties(Team(i)) + 1
End If
Select Case UnitMode(i)
Case Is = 0 'dead
GoTo ThisUnitIsDead
Case Is = 1 'run away
Call AIRun(i)
Case Is = 2 'attack
Call AIAttack(i)
End Select
Select Case UnitPos(i)
Case Is = 0 'North
XAdd = 0
YAdd = -1
Case Is = 1 'NorthEast
XAdd = 0.707
YAdd = -0.707
Case Is = 2 'East
XAdd = 1
YAdd = 0
Case Is = 3 'SouthEast
XAdd = 0.707
YAdd = 0.707
Case Is = 4 'South
XAdd = 0
YAdd = 1
Case Is = 5 'SouthWest
XAdd = -0.707
YAdd = 0.707
Case Is = 6 'West
XAdd = -1
YAdd = 0
Case Is = 7 'NorthWest
XAdd = -0.707
YAdd = -0.707
End Select
For X = 1 To Units
UnitAttacked = False
If Team(X) <> Team(i) And UnitMode(X) <> 0 Then 'if the other unit is on an opposing team then use weapons on it
For a = 0 To 6
If UnitY(X) >= UnitY(i) - Range(i, a) And _
UnitY(X) <= UnitY(i) + Range(i, a) Then
If UnitX(X) >= UnitX(i) - Range(i, a) And _
UnitX(X) <= UnitX(i) + Range(i, a) Then
HP(X) = Int(HP(X) - Firepower(i, a) * (1 - Armor(X, a) / 100))
UnitAttacked = True
End If
End If
Next a
End If
If UnitAttacked = True Then Attacked(i) = Attacked(i) + 1
If Attacked(i) = RateOfFire(i) Then Exit For  'if the unit has already attacked,
                                            'it can't attack again.
Next X
If Attacked(i) = 0 Then 'If the unit is not in combat then heal the unit
HP(i) = HP(i) + Healing(i)
End If
If HP(i) > Health(i) Then HP(i) = Health(i)
UnitX(i) = UnitX(i) + XAdd * Speed(i)
UnitY(i) = UnitY(i) + YAdd * Speed(i)
If UnitX(i) < 0 Then UnitX(i) = 0
If UnitX(i) > 14790 Then UnitX(i) = 14790
If UnitY(i) < 0 Then UnitY(i) = 0
If UnitY(i) > 10380 Then UnitY(i) = 10380
ThisUnitIsDead:
Next i
For i = 1 To Units
If UnitMode(i) = 0 Then 'dead
'BitBlt picField.hDC, UnitX(i), UnitY(i), picdeadMask.ScaleWidth, picdeadMask.ScaleHeight, picdeadMask.hDC, 0, 0, vbSrcAnd
'BitBlt picField.hDC, UnitX(i), UnitY(i), picDead(PColor).ScaleWidth, picDead(PColor).ScaleHeight, picDead(PColor).hDC, 0, 0, vbSrcPaint
End If
Next i
For i = 1 To Units
If UnitMode(i) <> 0 Then 'unit is alive
imgUnit(i).Move UnitX(i), UnitY(i)
End If
Next i
For i = 0 To 6
lblCasualties(i).Caption = Casualties(i)
Next i
lblTrackType = UnitName(TrackNo)
lblTrackHP = HP(TrackNo) & "/" & Health(TrackNo)
lblTrackMode = UnitMode(TrackNo)
lblTrackAttack = Attacked(TrackNo)
shpTrackHilite.Move UnitX(TrackNo), UnitY(TrackNo)
End Sub

Private Sub AIRun(Index)
Dim NW As Integer, NE As Integer, SE As Integer, SW As Integer, i
For i = 1 To Units
If Team(i) <> Team(Index) And UnitMode(i) <> 0 Then
If UnitX(i) < UnitX(Index) And UnitY(i) < UnitY(Index) Then NW = NW + 1
If UnitX(i) > UnitX(Index) And UnitY(i) < UnitY(Index) Then NE = NE + 1
If UnitX(i) > UnitX(Index) And UnitY(i) > UnitY(Index) Then SE = SE + 1
If UnitX(i) < UnitX(Index) And UnitY(i) > UnitY(Index) Then SW = SW + 1
End If
Next i
If NW > SW And NW > SE And NW > NE Then UnitPos(Index) = 3
If NE > SW And NE > SE And NE > NW Then UnitPos(Index) = 5
If SE > SW And SE > NW And SE > NE Then UnitPos(Index) = 7
If SW > NW And SW > SE And SW > NE Then UnitPos(Index) = 1
If NW = NE And NW + NE > SW + SE Then UnitPos(Index) = 4
If NE = SE And NE + SE > SW + NW Then UnitPos(Index) = 6
If SE = SW And SE + SW > NE + NW Then UnitPos(Index) = 0
If SW = NW And SW + NW > NE + SE Then UnitPos(Index) = 2
'If UnitPos(Index) = 7 And NW < NE + 2 And NW > NE - 2 Then UnitPos(i) = 0
'If UnitPos(Index) = 7 And NW < SW + 2 And NW > SW - 2 Then UnitPos(i) = 6
'If UnitPos(Index) = 5 And SW < NW + 2 And SW > NW - 2 Then UnitPos(i) = 6
'If UnitPos(Index) = 5 And SW < SE + 2 And SW > SE - 2 Then UnitPos(i) = 4
'If UnitPos(Index) = 3 And SE < SW + 2 And SE > SW - 2 Then UnitPos(i) = 4
'If UnitPos(Index) = 3 And SE < NE + 2 And SE > NE - 2 Then UnitPos(i) = 2
'If UnitPos(Index) = 1 And NE < SE + 2 And NE > SE - 2 Then UnitPos(i) = 2
'If UnitPos(Index) = 1 And NE < NW + 2 And NE > NW - 2 Then UnitPos(i) = 0




End Sub

Private Sub AIAttack(Index)
Dim NW As Integer, NE As Integer, SE As Integer, SW As Integer, i
For i = 1 To Units
If Team(i) <> Team(Index) And UnitMode(i) <> 0 Then
If UnitX(i) <= UnitX(Index) And UnitY(i) <= UnitY(Index) Then NW = NW + 1
If UnitX(i) >= UnitX(Index) And UnitY(i) <= UnitY(Index) Then NE = NE + 1
If UnitX(i) >= UnitX(Index) And UnitY(i) >= UnitY(Index) Then SE = SE + 1
If UnitX(i) <= UnitX(Index) And UnitY(i) >= UnitY(Index) Then SW = SW + 1
End If
Next i
If NW > SW And NW > SE And NW > NE Then UnitPos(Index) = 7
If NE > SW And NE > SE And NE > NW Then UnitPos(Index) = 1
If SE > SW And SE > NW And SE > NE Then UnitPos(Index) = 3
If SW > NW And SW > SE And SW > NE Then UnitPos(Index) = 5
If NW = NE And NW + NE > SW + SE Then UnitPos(Index) = 0
If NE = SE And NE + SE > SW + NW Then UnitPos(Index) = 2
If SE = SW And SE + SW > NE + NW Then UnitPos(Index) = 4
If SW = NW And SW + NW > NE + SE Then UnitPos(Index) = 6
If SW = NW And SE = NE And SW > SE Then UnitPos(Index) = 5
If SW = NW And SE = NE And SW < SE Then UnitPos(Index) = 3
'If UnitPos(Index) = 7 And NW < NE + 2 And NW > NE - 2 Then UnitPos(i) = 0
'If UnitPos(Index) = 7 And NW < SW + 2 And NW > SW - 2 Then UnitPos(i) = 6
'If UnitPos(Index) = 5 And SW < NW + 2 And SW > NW - 2 Then UnitPos(i) = 6
'If UnitPos(Index) = 5 And SW < SE + 2 And SW > SE - 2 Then UnitPos(i) = 4
'If UnitPos(Index) = 3 And SE < SW + 2 And SE > SW - 2 Then UnitPos(i) = 4
'If UnitPos(Index) = 3 And SE < NE + 2 And SE > NE - 2 Then UnitPos(i) = 2
'If UnitPos(Index) = 1 And NE < SE + 2 And NE > SE - 2 Then UnitPos(i) = 2
'If UnitPos(Index) = 1 And NE < NW + 2 And NE > NW - 2 Then UnitPos(i) = 0
End Sub
