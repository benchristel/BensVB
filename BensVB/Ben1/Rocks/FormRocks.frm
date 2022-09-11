VERSION 5.00
Begin VB.Form frmMine 
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Golddigger 1.0"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   15000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTime 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   2640
      Top             =   1740
   End
   Begin VB.CommandButton cmdKeypad 
      Caption         =   "9"
      Height          =   495
      Index           =   9
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdKeypad 
      Caption         =   "8"
      Height          =   495
      Index           =   8
      Left            =   10740
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdKeypad 
      Caption         =   "7"
      Height          =   495
      Index           =   7
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdKeypad 
      Caption         =   "6"
      Height          =   495
      Index           =   6
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4260
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdKeypad 
      Caption         =   "5"
      Height          =   495
      Index           =   5
      Left            =   10740
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4260
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdKeypad 
      Caption         =   "4"
      Height          =   495
      Index           =   4
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4260
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdKeypad 
      Caption         =   "3"
      Height          =   495
      Index           =   3
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdKeypad 
      Caption         =   "2"
      Height          =   495
      Index           =   2
      Left            =   10740
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdKeypad 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrIntro 
      Interval        =   5000
      Left            =   2640
      Top             =   720
   End
   Begin VB.Timer tmrExplode 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   3000
      Left            =   4560
      Top             =   120
   End
   Begin VB.Timer tmrExplode 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   3000
      Left            =   4080
      Top             =   120
   End
   Begin VB.Timer tmrExplode 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   3000
      Left            =   3600
      Top             =   120
   End
   Begin VB.Timer tmrExplode 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   3000
      Left            =   3120
      Top             =   120
   End
   Begin VB.Timer tmrExplode 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   3000
      Left            =   2640
      Top             =   120
   End
   Begin VB.TextBox txtMove 
      Height          =   375
      Left            =   -1000
      TabIndex        =   0
      Top             =   600
      Width           =   195
   End
   Begin VB.Image imgFood 
      Height          =   480
      Index           =   4
      Left            =   2640
      Picture         =   "FormRocks.frx":0000
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image imgFood 
      Height          =   480
      Index           =   3
      Left            =   2100
      Picture         =   "FormRocks.frx":08CA
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image imgFood 
      Height          =   480
      Index           =   2
      Left            =   1560
      Picture         =   "FormRocks.frx":1194
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image imgFood 
      Height          =   480
      Index           =   1
      Left            =   1020
      Picture         =   "FormRocks.frx":1A5E
      Top             =   4920
      Width           =   480
   End
   Begin VB.Image imgFood 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "FormRocks.frx":2328
      Top             =   4920
      Width           =   480
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      Caption         =   "May 1849"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   60
      Width           =   2535
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H000040C0&
      Caption         =   "Hello!  Welcome to Golddigger 1.0.  To skip this introduction, click in this box."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1395
      Left            =   3300
      TabIndex        =   2
      Top             =   3900
      Width           =   6795
   End
   Begin VB.Image imgGold 
      Height          =   480
      Index           =   3
      Left            =   2640
      Picture         =   "FormRocks.frx":2BF2
      Top             =   4380
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgGold 
      Height          =   480
      Index           =   2
      Left            =   2100
      Picture         =   "FormRocks.frx":34BC
      Top             =   4380
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarryRock 
      Height          =   480
      Left            =   1020
      Picture         =   "FormRocks.frx":3D86
      Top             =   4380
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCarryWood 
      Height          =   480
      Left            =   480
      Picture         =   "FormRocks.frx":4650
      Top             =   4380
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblMoney 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "$1,000.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   5460
      Width           =   2115
   End
   Begin VB.Image imgGold 
      Height          =   480
      Index           =   1
      Left            =   1560
      Picture         =   "FormRocks.frx":4F1A
      Top             =   4380
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTNT 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "FormRocks.frx":57E4
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgScaffold 
      Height          =   480
      Left            =   600
      Picture         =   "FormRocks.frx":60AE
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgTunnel 
      Height          =   480
      Left            =   60
      Picture         =   "FormRocks.frx":6978
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPlayer 
      Height          =   480
      Left            =   1680
      Picture         =   "FormRocks.frx":7242
      Top             =   60
      Width           =   480
   End
   Begin VB.Image imgForest 
      Height          =   480
      Left            =   1140
      Picture         =   "FormRocks.frx":754C
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSky 
      Height          =   480
      Left            =   600
      Picture         =   "FormRocks.frx":7E16
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSpace 
      Height          =   480
      Index           =   0
      Left            =   60
      Picture         =   "FormRocks.frx":86E0
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Shape shpGoldFlash 
      BorderColor     =   &H000080FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1500
      Top             =   4380
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Shape shpInvFlash 
      BorderColor     =   &H000080FF&
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   2115
      Left            =   420
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "frmMine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
Dim i
Randomize
For i = 1 To 140
Load imgSpace(i)
Select Case i
Case Is < 21
imgSpace(i).Top = 480
imgSpace(i).Left = 480 * i
Case Is < 41
imgSpace(i).Top = 960
imgSpace(i).Left = 480 * (i - 20)
Case Is < 61
imgSpace(i).Top = 1440
imgSpace(i).Left = 480 * (i - 40)
Case Is < 81
imgSpace(i).Top = 1920
imgSpace(i).Left = 480 * (i - 60)
Case Is < 101
imgSpace(i).Top = 2400
imgSpace(i).Left = 480 * (i - 80)
Case Is < 121
imgSpace(i).Top = 2880
imgSpace(i).Left = 480 * (i - 100)
Case Else
imgSpace(i).Top = 3360
imgSpace(i).Left = 480 * (i - 120)
End Select
imgSpace(i).Visible = True
SpaceState(i) = "Rock"
Next i
For i = 1 To 25
imgSpace(i).Picture = imgSky.Picture
SpaceState(i) = "Sky"
Next i
For i = 28 To 40
imgSpace(i).Picture = imgForest.Picture
SpaceState(i) = "Forest"
Next i
For i = 41 To 43
imgSpace(i).Picture = imgSky.Picture
SpaceState(i) = "Sky"
Next i
For i = 1 To 5
Load imgTNT(i)
imgTNT(i).Top = 3780
imgTNT(i).Left = 480 * i
imgTNT(i).Visible = True
imgTNT(i).ZOrder
Next i
For i = 1 To 3
GoldPos(i) = Int(Rnd * 80 + 61)
Next i
imgPlayer.Top = 1440
imgPlayer.Left = 480
Position = 41
CarryWood = False
CarryRock = False
TNTPlaced = 1
Food = 5
Money = 1000
lblMoney.Caption = FormatCurrency(Money)
Intro = True
Week = 1
WeekNo = 1
MonthNo = 5
Year = 1849
End Sub

Private Sub lblInfo_Click()
Dim i
    If Intro = True Then
    Intro = False
    tmrIntro.Enabled = False
    lblInfo.Caption = "CLICK HERE FOR HELP"
    For i = 1 To 3
    imgGold(i).Visible = False
    Next i
    shpGoldFlash.Visible = False
    shpInvFlash.Visible = False
    imgCarryRock.Visible = False
    imgCarryWood.Visible = False
    For i = 1 To 9
    cmdKeypad(i).Visible = False
    Next i
    tmrTime.Enabled = True
    Else
    lblInfo.Caption = "HELP SYSTEM ACTIVATED" & vbLf & "CLICK HERE TO CLOSE"
    Help = True
    Exit Sub
    End If
    If Help = True Then
    Help = False
    End If
End Sub

Private Sub tmrExplode_Timer(Index As Integer)
Dim i, n
imgSpace(TNTpos(Index)).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index)) = "Sky"
If imgTNT(Index).Top > 480 And imgTNT(Index).Left > 480 Then
imgSpace(TNTpos(Index) - 21).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) - 21) = "Sky"
imgSpace(TNTpos(Index) - 20).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) - 20) = "Sky"
imgSpace(TNTpos(Index) - 1).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) - 1) = "Sky"
End If
If imgTNT(Index).Left > 480 And imgTNT(Index).Top < 3360 Then
imgSpace(TNTpos(Index) + 19).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) + 19) = "Sky"
imgSpace(TNTpos(Index) + 20).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) + 20) = "Sky"
imgSpace(TNTpos(Index) + 1).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) + 1) = "Sky"
End If
If imgTNT(Index).Left < 9600 And imgTNT(Index).Top > 480 Then
imgSpace(TNTpos(Index) - 19).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) - 19) = "Sky"
imgSpace(TNTpos(Index) + 1).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) + 1) = "Sky"
imgSpace(TNTpos(Index) - 20).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) - 20) = "Sky"
End If
If imgTNT(Index).Left < 9600 And imgTNT(Index).Top < 3360 Then
imgSpace(TNTpos(Index) + 21).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) + 21) = "Sky"
imgSpace(TNTpos(Index) + 20).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) + 20) = "Sky"
imgSpace(TNTpos(Index) + 1).Picture = imgTunnel.Picture
SpaceState(TNTpos(Index) + 1) = "Sky"
End If
Unload imgTNT(Index)
tmrExplode(Index).Enabled = False
Avalanche
PlayerFall
For n = 61 To 140
If SpaceState(n) = "Sky" Then
    For i = 1 To 3
    If GoldPos(i) = n And imgGold(i).Visible = False Then
    imgGold(i).Visible = True
    lblMoney.Caption = FormatCurrency(Money)
    End If
    Next i
    End If
Next n
    If imgGold(1).Visible = True And imgGold(2).Visible = True And imgGold(3).Visible = True Then
    Game = "Won"
    lblInfo.Caption = "You have found $3000 worth of gold in your mine.  You return home and live happily ever after.  You have won.  Your total money is " _
    & FormatCurrency(Money + 3000) & "."
    End If

End Sub

Private Sub tmrIntro_Timer()
Dim i
IntroCount = IntroCount + 1
Select Case IntroCount
Case Is = 1
lblInfo.Caption = "The object of the game is to accumulate 3 gold bars without running out of money or getting killed."
shpGoldFlash.Visible = True
For i = 1 To 3
imgGold(i).Visible = True
Next i
Case Is = 2
lblInfo.Caption = "This is the Mining Window, where you'll be doing most of your work."
shpGoldFlash.Visible = False
For i = 1 To 3
imgGold(i).Visible = False
Next i
Case Is = 3
lblInfo.Caption = "To the left is your Inventory.  It shows everything you are carrying at the moment."
shpInvFlash.Visible = True
Case Is = 4
lblInfo.Caption = "You can carry wood, stone, money, dynamite, and up to 3 bars of gold."
imgCarryWood.Visible = True
imgCarryRock.Visible = True
For i = 1 To 3
imgGold(i).Visible = True
Next i
Case Is = 5
lblInfo.Caption = "You move around the Mining Window by pressing keys on the keypad."
For i = 1 To 9
cmdKeypad(i).Visible = True
cmdKeypad(i).BackColor = RGB(200, 200, 200)
Next i
imgCarryWood.Visible = False
imgCarryRock.Visible = False
For i = 1 To 3
imgGold(i).Visible = False
shpInvFlash.Visible = False
Next i
Case Is = 6
lblInfo.Caption = "Use the 1 and 3 keys to move left and right across the terrain."
cmdKeypad(1).BackColor = RGB(255, 255, 0)
cmdKeypad(3).BackColor = RGB(255, 255, 0)
Case Is = 7
lblInfo.Caption = "You can move up one level on the terrain by jumping with the 7 and 9 keys."
cmdKeypad(1).BackColor = RGB(200, 200, 200)
cmdKeypad(3).BackColor = RGB(200, 200, 200)
cmdKeypad(7).BackColor = RGB(255, 255, 0)
cmdKeypad(9).BackColor = RGB(255, 255, 0)
Case Is = 8
lblInfo.Caption = "Keys 8 and 2 let you dig up and down, and also let you move vertically on scaffolding."
cmdKeypad(7).BackColor = RGB(200, 200, 200)
cmdKeypad(9).BackColor = RGB(200, 200, 200)
cmdKeypad(8).BackColor = RGB(255, 255, 0)
cmdKeypad(2).BackColor = RGB(255, 255, 0)
Case Is = 9
lblInfo.Caption = "The 4 and 6 keys are strictly for digging left and right."
cmdKeypad(8).BackColor = RGB(200, 200, 200)
cmdKeypad(2).BackColor = RGB(200, 200, 200)
cmdKeypad(4).BackColor = RGB(255, 255, 0)
cmdKeypad(6).BackColor = RGB(255, 255, 0)
Case Is = 10
lblInfo.Caption = "You can also cut down trees or demolish scaffolding by pressing the 5 key."
cmdKeypad(4).BackColor = RGB(200, 200, 200)
cmdKeypad(6).BackColor = RGB(200, 200, 200)
cmdKeypad(5).BackColor = RGB(255, 255, 0)
Case Is = 11
lblInfo.Caption = "If you're carrying rock in your inventory, you can dump it by pressing the D key."
cmdKeypad(5).BackColor = RGB(200, 200, 200)
Case Is = 12
lblInfo.Caption = "Likewise, you can build scaffolding if you're carrying wood by pressing the B key."
Case Is = 13
lblInfo.Caption = "Scaffolding allows you to travel up and down in deep shafts, which is useful when mining."
Case Is = 14
lblInfo.Caption = "You can drop dynamite by pressing the ENTER key.  You can remove large areas of rock this way."
Case Is = 15
lblInfo.Caption = "Another very important feature is the Mining Town Window, which you can access by pressing the HOME key when above ground."
Case Is = 16
lblInfo.Caption = "This allows you to buy wood and other goods, and borrow and loan money."
Case Is = 17
lblInfo.Caption = "If you still don't get it, or you just need a helpful hint, click here for help."
For i = 1 To 9
cmdKeypad(i).Visible = False
Next i
Case Is = 18
lblInfo.Caption = "CLICK HERE FOR HELP"
tmrTime.Enabled = True
tmrIntro.Enabled = False
Intro = False
End Select
End Sub

Private Sub tmrTime_Timer()
Dim i
If Game = "Lost" Or Game = "Won" Then
tmrTime.Enabled = False
Exit Sub
End If
 Week = Week + 1
 WeekNo = WeekNo + 1
 If WeekNo = 5 Then
 MonthNo = MonthNo + 1
 WeekNo = 1
 End If
 If MonthNo = 13 Then
 MonthNo = 1
 Year = Year + 1
 End If
 Select Case MonthNo
 Case Is = 1
 Month = "January"
  Case Is = 2
 Month = "February"
  Case Is = 3
 Month = "March"
  Case Is = 4
 Month = "April"
  Case Is = 5
 Month = "May"
  Case Is = 6
 Month = "June"
  Case Is = 7
 Month = "July"
  Case Is = 8
 Month = "August"
  Case Is = 9
 Month = "September"
  Case Is = 10
 Month = "October"
  Case Is = 11
 Month = "November"
  Case Is = 12
 Month = "December"
 End Select
 lblDate.Caption = Month & " " & Year
 WoodPrice = 9 + (1.008 ^ Week)
TNTPrice = 49 + (1.03 ^ Week)
FoodPrice = 3 + (1.003 ^ Week)
 If Week Mod 2 = 0 Then
 If Food > 0 Then
 Food = Food - 1
 Else
 Food = 5
 Money = Money - 0.4
 Money = Money - FoodPrice * 5
 lblMoney.Caption = FormatCurrency(Money)
' Game = "Lost"
' lblInfo.Caption = "You have starved to death because you forgot to buy food.  Others have been warned not to imitate this fatal breech of nutritional intake practices."
 End If
  For i = 0 To 4
 imgFood(i).Visible = False
 Next i
 For i = 0 To Food - 1
 imgFood(i).Visible = True
 Next i
End If
End Sub

Private Sub txtMove_KeyDown(KeyCode As Integer, Shift As Integer)
Dim falling As Boolean, i, n

If Game = "Lost" Or Game = "Won" Or Intro = True Then Exit Sub
Select Case KeyCode
Case Is = 97 ' move left
    If imgPlayer.Left > 480 Then
        If SpaceState(Position - 1) = "Sky" Or _
        SpaceState(Position - 1) = "Forest" Or _
        SpaceState(Position - 1) = "Scaffold" Then
        imgPlayer.Left = imgPlayer.Left - 480
        Avalanche
        Position = Position - 1
        Money = Money - 0.25
        End If
    End If
Case Is = 99 ' move right
    If imgPlayer.Left < 9600 Then
        If SpaceState(Position + 1) = "Sky" Or _
        SpaceState(Position + 1) = "Forest" Or _
        SpaceState(Position + 1) = "Scaffold" Then
        imgPlayer.Left = imgPlayer.Left + 480
        Avalanche
        Position = Position + 1
        Money = Money - 0.25
        End If
    End If
Case Is = 103 ' jump left
    If imgPlayer.Left > 480 And imgPlayer.Top > 480 Then
        If SpaceState(Position - 21) = "Sky" Or _
        SpaceState(Position - 21) = "Forest" Then
            If SpaceState(Position - 20) = "Sky" Or _
            SpaceState(Position - 20) = "Forest" Then
            imgPlayer.Left = imgPlayer.Left - 480
            imgPlayer.Top = imgPlayer.Top - 480
            Position = Position - 21
            Money = Money - 0.25
            End If
        End If
    End If
Case Is = 105 ' jump right
If imgPlayer.Left < 9600 And imgPlayer.Top > 480 Then
    If SpaceState(Position - 19) = "Sky" Or _
    SpaceState(Position - 19) = "Forest" Then
        If SpaceState(Position - 20) = "Sky" Or _
        SpaceState(Position - 20) = "Forest" Then
        imgPlayer.Left = imgPlayer.Left + 480
        imgPlayer.Top = imgPlayer.Top - 480
        Position = Position - 19
        Money = Money - 0.25
        End If
    End If
End If
Case Is = 98 ' dig down
    If imgPlayer.Top < 3360 Then
    If CarryRock = False And SpaceState(Position + 20) = "Rock" Then
    imgSpace(Position + 20).Picture = imgTunnel.Picture
    SpaceState(Position + 20) = "Sky"
    imgPlayer.Top = imgPlayer.Top + 480
    Position = Position + 20
    CarryRock = True
    Money = Money - 2
    For i = 1 To 3
    If GoldPos(i) = Position Then
    imgGold(i).Visible = True
    End If
    Next i
    End If
    End If
    If imgPlayer.Top < 3360 Then
    If SpaceState(Position + 20) = "Scaffold" Then
    imgPlayer.Top = imgPlayer.Top + 480
    Position = Position + 20
    Money = Money - 0.25
    End If
    End If
Case Is = 100 ' dig left
If imgPlayer.Left > 480 And CarryRock = False And SpaceState(Position - 1) = "Rock" Then
imgSpace(Position - 1).Picture = imgTunnel.Picture
SpaceState(Position - 1) = "Sky"
imgPlayer.Left = imgPlayer.Left - 480
CarryRock = True
Position = Position - 1
Money = Money - 2
    For i = 1 To 3
    If GoldPos(i) = Position Then
    imgGold(i).Visible = True
    End If
    Next i
End If
Case Is = 102 ' dig right
If imgPlayer.Left < 9600 And CarryRock = False And SpaceState(Position + 1) = "Rock" Then
imgSpace(Position + 1).Picture = imgTunnel.Picture
SpaceState(Position + 1) = "Sky"
imgPlayer.Left = imgPlayer.Left + 480
CarryRock = True
Position = Position + 1
Money = Money - 2
    For i = 1 To 3
    If GoldPos(i) = Position Then
    imgGold(i).Visible = True
    End If
    Next i
End If
Case Is = 104 ' move up
If imgPlayer.Top > 480 And SpaceState(Position) = "Scaffold" Then
imgPlayer.Top = imgPlayer.Top - 480
Position = Position - 20
If SpaceState(Position) = "Rock" Then
imgSpace(Position) = imgTunnel.Picture
Money = Money - 1.75
    For i = 1 To 3
    If GoldPos(i) = Position Then
    imgGold(i).Visible = True
    End If
    Next i
End If
Money = Money - 0.25
End If
Case Is = 68 ' dump rock
If CarryRock = True Then
imgSpace(Position).Picture = imgSpace(0).Picture
SpaceState(Position) = "Rock"
CarryRock = False
If Position > 20 Then
If SpaceState(Position - 20) = "Rock" Then
Game = "Lost"
lblInfo.Caption = "You have been buried in a cave-in caused by your own folly.  You become a lovely fossil and are put on display at a natural history museum 10,000 years later.  You have lost."
End If
End If
Money = Money - 0.75
End If
Case Is = 66 ' build scaffolding
If CarryWood = True And SpaceState(Position) = "Sky" Then
imgSpace(Position).Picture = imgScaffold.Picture
SpaceState(Position) = "Scaffold"
CarryWood = False
Money = Money - 3.1
End If
Case Is = 101 ' cut trees
If CarryWood = False Then
If SpaceState(Position) = "Forest" Then
imgSpace(Position).Picture = imgSky.Picture
SpaceState(Position) = "Sky"
CarryWood = True
Money = Money - 3
End If
If SpaceState(Position) = "Scaffold" Then
imgSpace(Position).Picture = imgTunnel.Picture
SpaceState(Position) = "Sky"
CarryWood = True
Money = Money - 3
End If
End If
Case Is = 13 ' drop dynamite
If TNTPlaced < 6 Then
imgTNT(TNTPlaced).ZOrder
imgTNT(TNTPlaced).Top = imgSpace(Position).Top
imgTNT(TNTPlaced).Left = imgSpace(Position).Left
TNTpos(TNTPlaced) = Position
tmrExplode(TNTPlaced).Enabled = True
TNTPlaced = TNTPlaced + 1
Money = Money - 12
End If
Case Is = 36 ' go to mining town
If imgSpace(Position).Picture = imgSky.Picture Or imgSpace(Position).Picture = imgForest.Picture Then
InTown = True
Money = Money - 0.4
tmrTime.Enabled = False
frmTown.Show 1
End If
End Select
falling = True
    Do Until falling = False
        If imgPlayer.Top < 3360 Then
        If SpaceState(Position) <> "Scaffold" And SpaceState(Position + 20) <> "Scaffold" Then
            If SpaceState(Position + 20) = "Sky" Or SpaceState(Position + 20) = "Forest" Then
            imgPlayer.Top = imgPlayer.Top + 480
            Position = Position + 20
            Else
            falling = False
            End If
            Else
            falling = False
        End If
        Else
        falling = False
        End If
    Loop
        If Money <= 0 Then
    Game = "Lost"
    Money = 0
    End If
    lblMoney.Caption = FormatCurrency(Money)
    If CarryWood = True Then
    imgCarryWood.Visible = True
    Else
    imgCarryWood.Visible = False
    End If
    If CarryRock = True Then
    imgCarryRock.Visible = True
    Else
    imgCarryRock.Visible = False
    End If
    For i = 1 To 40
    If imgSpace(i).Picture = imgTunnel.Picture Then
    imgSpace(i).Picture = imgSky.Picture
    End If
    Next i
    If Money <= 0 Then
    Game = "Lost"
    Money = 0
    lblInfo.Caption = "After several years of mining, you run out of money.  You are unable to return home and so you become a clothespin salesman in the mining town.  You have lost."
    End If
    If imgGold(1).Visible = True And imgGold(2).Visible = True And imgGold(3).Visible = True Then
    Game = "Won"
    lblInfo.Caption = "You have found $3000 worth of gold in your mine.  You return home and live happily ever after.  You have won.  Your total money is " _
    & FormatCurrency(Money) & "."
    End If
    'MsgBox KeyCode
End Sub


Private Sub Avalanche()
Dim i, count
Do Until count = 3
For i = 121 To 140
If SpaceState(i) = "Sky" And i > 20 Then
If SpaceState(i - 20) = "Rock" Then
SpaceState(i - 20) = "Sky"
SpaceState(i) = "Rock"
imgSpace(i).Picture = imgSpace(0).Picture
imgSpace(i - 20).Picture = imgTunnel.Picture
End If
End If
Next i
For i = 101 To 120
If SpaceState(i) = "Sky" And i > 20 Then
If SpaceState(i - 20) = "Rock" Then
SpaceState(i - 20) = "Sky"
SpaceState(i) = "Rock"
imgSpace(i).Picture = imgSpace(0).Picture
imgSpace(i - 20).Picture = imgTunnel.Picture
End If
End If
Next i
For i = 81 To 100
If SpaceState(i) = "Sky" And i > 20 Then
If SpaceState(i - 20) = "Rock" Then
SpaceState(i - 20) = "Sky"
SpaceState(i) = "Rock"
imgSpace(i).Picture = imgSpace(0).Picture
imgSpace(i - 20).Picture = imgTunnel.Picture
End If
End If
Next i
For i = 61 To 80
If SpaceState(i) = "Sky" And i > 20 Then
If SpaceState(i - 20) = "Rock" Then
SpaceState(i - 20) = "Sky"
SpaceState(i) = "Rock"
imgSpace(i).Picture = imgSpace(0).Picture
imgSpace(i - 20).Picture = imgTunnel.Picture
End If
End If
Next i
For i = 41 To 60
If SpaceState(i) = "Sky" And i > 20 Then
If SpaceState(i - 20) = "Rock" Then
SpaceState(i - 20) = "Sky"
SpaceState(i) = "Rock"
imgSpace(i).Picture = imgSpace(0).Picture
imgSpace(i - 20).Picture = imgTunnel.Picture
End If
End If
Next i
For i = 21 To 40
If SpaceState(i) = "Sky" And i > 20 Then
If SpaceState(i - 20) = "Rock" Then
SpaceState(i - 20) = "Sky"
SpaceState(i) = "Rock"
imgSpace(i).Picture = imgSpace(0).Picture
imgSpace(i - 20).Picture = imgTunnel.Picture
End If
End If
Next i
count = count + 1
Loop
    For i = 1 To 40
    If imgSpace(i).Picture = imgTunnel.Picture Then
    imgSpace(i).Picture = imgSky.Picture
    End If
    Next i
End Sub

Private Sub PlayerFall()
Dim falling
falling = True
    Do Until falling = False
        If imgPlayer.Top < 3360 Then
            If SpaceState(Position + 20) = "Sky" Or SpaceState(Position + 20) = "Forest" Then
            imgPlayer.Top = imgPlayer.Top + 480
            Position = Position + 20
            Else
            falling = False
            End If
            Else
            falling = False
        End If
    Loop

End Sub



