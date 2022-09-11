VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Space Shuttle Launch Adventure"
   ClientHeight    =   7710
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   11805
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPayloadSpecial 
      Height          =   285
      Left            =   1800
      TabIndex        =   23
      Top             =   7440
      Width           =   1815
   End
   Begin VB.TextBox txtMisSpecial 
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      Top             =   7080
      Width           =   1815
   End
   Begin VB.TextBox txtCoPilot 
      Height          =   285
      Left            =   1800
      TabIndex        =   21
      Top             =   6720
      Width           =   1815
   End
   Begin VB.TextBox txtPilot 
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox txtCommander 
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Timer tmrLiftoff 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5280
      Top             =   6360
   End
   Begin VB.ComboBox cboOrbiters 
      Height          =   315
      ItemData        =   "SHUTTLE.frx":0000
      Left            =   1800
      List            =   "SHUTTLE.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox txtCountdown 
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Timer tmrCountdown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5280
      Top             =   5880
   End
   Begin VB.CommandButton cmdLiftoff 
      Caption         =   "Liftoff!"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdOrbit 
      Caption         =   "Enter Orbit"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4080
      TabIndex        =   12
      Top             =   6840
      Width           =   1000
   End
   Begin VB.CommandButton cmdJettF 
      Caption         =   "Jettison Fuel Tank"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4080
      TabIndex        =   11
      Top             =   6360
      Width           =   1000
   End
   Begin VB.CommandButton cmdJettB 
      Caption         =   "Jettison Boosters"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4080
      TabIndex        =   10
      Top             =   5880
      Width           =   1000
   End
   Begin VB.CommandButton cmdBoosters3 
      Caption         =   "Boosters"
      Height          =   400
      Left            =   8280
      TabIndex        =   8
      Top             =   840
      Width           =   1075
   End
   Begin VB.CommandButton cmdBoosters2 
      Caption         =   "Boosters"
      Height          =   400
      Left            =   5160
      TabIndex        =   7
      Top             =   840
      Width           =   1075
   End
   Begin VB.CommandButton cmdBoosters1 
      Caption         =   "Boosters"
      Height          =   400
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   1075
   End
   Begin VB.CommandButton cmdZoom3 
      Caption         =   "Zoom Out"
      Height          =   400
      Left            =   7320
      TabIndex        =   5
      Top             =   840
      Width           =   1000
   End
   Begin VB.CommandButton cmdZoom2 
      Caption         =   "Zoom Out"
      Height          =   400
      Left            =   4200
      TabIndex        =   4
      Top             =   840
      Width           =   1000
   End
   Begin VB.CommandButton cmdZoom1 
      Caption         =   "Zoom Out"
      Height          =   400
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   1000
   End
   Begin VB.CommandButton cmdShuttle3 
      Caption         =   "Shuttle"
      Height          =   400
      Left            =   6360
      TabIndex        =   2
      Top             =   840
      Width           =   1000
   End
   Begin VB.CommandButton cmdShuttle2 
      Caption         =   "Shuttle"
      Height          =   400
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   1000
   End
   Begin VB.CommandButton cmdShuttle1 
      Caption         =   "Shuttle"
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1000
   End
   Begin VB.Label lblCoPilot 
      Alignment       =   1  'Right Justify
      Caption         =   "Co Pilot's name:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblPilot 
      Alignment       =   1  'Right Justify
      Caption         =   "Pilot's name:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label lblCommander 
      Alignment       =   1  'Right Justify
      Caption         =   "Commander's name:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lblMission 
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   11655
   End
   Begin VB.Label Label1 
      Caption         =   "Choose an orbiter here"
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label lblInfo 
      Caption         =   "Type the length of the countdown here "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   5400
      Width           =   3495
   End
   Begin VB.Label lblEvents 
      Alignment       =   2  'Center
      Caption         =   "Currently Refueling..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   9255
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   3000
      Left            =   6360
      Picture         =   "SHUTTLE.frx":0004
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   3000
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3000
      Left            =   3240
      Picture         =   "SHUTTLE.frx":A2BE
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   3000
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3000
      Left            =   120
      Picture         =   "SHUTTLE.frx":14578
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   3000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear mission list"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuView 
         Caption         =   "View Mission List"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuMission 
      Caption         =   "Mission"
      Begin VB.Menu mnuLiftoff 
         Caption         =   "Liftoff"
         Enabled         =   0   'False
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuJettB 
         Caption         =   "Jettison Boosters"
         Enabled         =   0   'False
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuJettF 
         Caption         =   "Jettison Fuel Tank"
         Enabled         =   0   'False
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEnterOrbit 
         Caption         =   "Enter Orbit"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuAbort 
         Caption         =   "Abort Mission"
         Enabled         =   0   'False
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuPurpose 
         Caption         =   "Mission Purpose..."
         Begin VB.Menu mnuP_Deploy 
            Caption         =   "Deploy Satillite..."
            Begin VB.Menu mnuMilitary_D 
               Caption         =   "Military"
               Shortcut        =   {F2}
            End
            Begin VB.Menu mnuComm_D 
               Caption         =   "Communications"
               Shortcut        =   {F3}
            End
         End
         Begin VB.Menu mnuP_Retrieve 
            Caption         =   "Retreive Satellite..."
            Begin VB.Menu mnuMilitary_R 
               Caption         =   "Military"
               Shortcut        =   ^{F2}
            End
            Begin VB.Menu mnuComm_R 
               Caption         =   "Communications"
               Shortcut        =   ^{F3}
            End
         End
         Begin VB.Menu mnuFindStar 
            Caption         =   "Find New Star"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuTrain 
            Caption         =   "Training "
            Shortcut        =   {F5}
         End
      End
   End
   Begin VB.Menu mnuOrbit 
      Caption         =   "Orbit"
      Begin VB.Menu mnuDeploy 
         Caption         =   "Deploy Satellite"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu mnuWalk 
         Caption         =   "Retreive Satellite/Spacewalk"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuLand 
         Caption         =   "Land Shuttle"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuCheat 
      Caption         =   "Cheat"
      Begin VB.Menu mnuEnable 
         Caption         =   "Enable Cheat"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuAuto_L 
         Caption         =   "Auto-land"
         Enabled         =   0   'False
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuAuto_G 
         Caption         =   "Auto-grab"
         Enabled         =   0   'False
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuWatch 
         Caption         =   "Watch Mode "
         Enabled         =   0   'False
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Countdown
Dim MisPurpose
Dim MissionNumber
Dim ShuttleName
Dim FileName
Dim objPic1 As Picture
Dim Commander, Pilot, CoPilot, MisSpecial, PayloadSpecial, MisInfo As InfoType

Private Sub cboOrbiters_Change()
If cboOrbiters.Text = "Enterprise" Then
MisInfo.Shuttle = "Enterprise"
'ShuttleName = "Enterprise"
End If
If cboOrbiters.Text = "Columbia" Then
ShuttleName = "Columbia"
End If
If cboOrbiters.Text = "Challenger" Then
ShuttleName = "Challenger"
End If
If cboOrbiters.Text = "Atlantis" Then
ShuttleName = "Atlantis"
End If
If cboOrbiters.Text = "Discovery" Then
ShuttleName = "Discovery"
End If
If Val(txtCountdown.Text) > 0 And cboOrbiters.Text <> "" And MisPurpose <> "" Then
cmdLiftoff.Enabled = True
End If
If Val(txtCountdown.Text) < 1 Or txtCountdown.Text = "" Or cboOrbiters.Text = "" Or MisPurpose = "" Then
cmdLiftoff.Enabled = False
End If
Countdown = txtCountdown.Text

End Sub



Private Sub cmdJettB_Click()
cmdJettF.Enabled = True
cmdJettB.Enabled = False
End Sub

Private Sub cmdJettF_Click()
cmdOrbit.Enabled = True
cmdJettF.Enabled = False
mnuOrbit.Enabled = True
End Sub

Private Sub cmdLiftoff_Click()
tmrCountdown.Enabled = True
lblEvents.Caption = "T-minus " & Countdown & " seconds and counting"
cmdLiftoff.Enabled = False
lblInfo.Visible = False
txtCountdown.Visible = False
cboOrbiters.Visible = False
Label1.Visible = False
mnuAbort.Enabled = True
MisInfo = "STS-" & MissionNumber & ": " & cboOrbiters.Text & _
  ";" & MisPurpose & ";Commander: " & Commander & ", Pilot: " & Pilot & ", Co-Pilot: " & CoPilot & ", Mission Specialist: " & MisSpecial & ", Payload Specialist: " & PayloadSpecial & "; countdown started at " & Time()
lblMission.Caption = MisInfo
Form2.List1.AddItem MisInfo
'  Image1.Picture = App.Path & "\launch.jpg"
'  Image3.Picture = "launch.jpg"
'    Image2.Picture = "launch.jpg"

End Sub



Private Sub Form_Load()
cboOrbiters.AddItem ("Endeavor")
If MissionNumber = 5 Then
cboOrbiters.AddItem ("Columbia")
End If
If MissionNumber = 10 Then
cboOrbiters.AddItem ("Challenger")
End If
If MissionNumber = 15 Then
cboOrbiters.AddItem ("Atlantis")
End If
If MissionNumber = 20 Then
cboOrbiters.AddItem ("Discovery")
End If
mnuOrbit.Enabled = False
FileName = App.Path & "\MsnNumber.txt"
Open FileName For Input As #1
Input #1, MissionNumber
Close #1
End Sub

Private Sub mnuAbort_Click()
tmrCountdown.Enabled = False
Countdown = 0
lblEvents.Caption = "Mission Aborted"
End Sub

Private Sub mnuClear_Click()
MissionNumber = 1
Open FileName For Output As #1
Print #1, MissionNumber
Close #1
End Sub

Private Sub mnuComm_D_Click()
MisPurpose = "Deploy Communications Satellite"
If Val(txtCountdown.Text) > 0 And cboOrbiters.Text <> "" And MisPurpose <> "" Then
cmdLiftoff.Enabled = True
End If
If Val(txtCountdown.Text) < 1 Or txtCountdown.Text = "" Or cboOrbiters.Text = "" Or MisPurpose = "" Then
cmdLiftoff.Enabled = False
End If
Countdown = txtCountdown.Text

End Sub

Private Sub mnuComm_R_Click()
MisPurpose = "Retrieve communications satillite"
If Val(txtCountdown.Text) > 0 And cboOrbiters.Text <> "" And MisPurpose <> "" Then
cmdLiftoff.Enabled = True
End If
If Val(txtCountdown.Text) < 1 Or txtCountdown.Text = "" Or cboOrbiters.Text = "" Or MisPurpose = "" Then
cmdLiftoff.Enabled = False
End If
Countdown = txtCountdown.Text

End Sub

Private Sub mnuEnable_Click()
Dim Result
 Result = MsgBox("Really enable Cheat Mode?", vbQuestion + vbYesNo, "Cheat Mode")
 'MsgBox Result   'Yes=6 and No = 7
 If Result = 6 Then
  mnuAuto_L.Enabled = True
 mnuAuto_G.Enabled = True
 mnuWatch.Enabled = True
 mnuEnable.Enabled = False
 End If
End Sub

Private Sub mnuFindStar_Click()
MisPurpose = "Find new star"
If Val(txtCountdown.Text) > 0 And cboOrbiters.Text <> "" And MisPurpose <> "" Then
cmdLiftoff.Enabled = True
End If
If Val(txtCountdown.Text) < 1 Or txtCountdown.Text = "" Or cboOrbiters.Text = "" Or MisPurpose = "" Then
cmdLiftoff.Enabled = False
End If
Countdown = txtCountdown.Text

End Sub

Private Sub mnuMilitary_D_Click()
MisPurpose = "Deploy Spy Satellite"
If Val(txtCountdown.Text) > 0 And cboOrbiters.Text <> "" And MisPurpose <> "" Then
cmdLiftoff.Enabled = True
End If
If Val(txtCountdown.Text) < 1 Or txtCountdown.Text = "" Or cboOrbiters.Text = "" Or MisPurpose = "" Then
cmdLiftoff.Enabled = False
End If
Countdown = txtCountdown.Text

End Sub

Private Sub mnuMilitary_R_Click()
MisPurpose = "Retrieve spy satellite"
If Val(txtCountdown.Text) > 0 And cboOrbiters.Text <> "" And MisPurpose <> "" Then
cmdLiftoff.Enabled = True
End If
If Val(txtCountdown.Text) < 1 Or txtCountdown.Text = "" Or cboOrbiters.Text = "" Or MisPurpose = "" Then
cmdLiftoff.Enabled = False
End If
Countdown = txtCountdown.Text

End Sub

Private Sub mnuQuit_Click()
End
End Sub


Private Sub mnuTrain_Click()
MisPurpose = "Train Astronauts"
If Val(txtCountdown.Text) > 0 And cboOrbiters.Text <> "" And MisPurpose <> "" Then
cmdLiftoff.Enabled = True
End If
If Val(txtCountdown.Text) < 1 Or txtCountdown.Text = "" Or cboOrbiters.Text = "" Or MisPurpose = "" Then
cmdLiftoff.Enabled = False
End If
Countdown = txtCountdown.Text

End Sub

Private Sub mnuView_Click()
Form2.Visible = True
Form1.Visible = False
End Sub

Private Sub tmrCountdown_Timer()
Countdown = Countdown - 1
lblEvents.Caption = "T-minus " & Countdown & " seconds and counting"
If Countdown = 0 Then
lblEvents.Caption = "Liftoff!"
Set objPic1 = LoadPicture(App.Path & "\launch.jpg")
Image1.Picture = objPic1
Set objPic1 = LoadPicture(App.Path & "\launch.jpg")
Image2.Picture = objPic1
Set objPic1 = LoadPicture(App.Path & "\launch.jpg")
Image3.Picture = objPic1
tmrCountdown.Enabled = False
cmdLiftoff.Enabled = False
cmdJettB.Enabled = True
Open FileName For Output As #1
Print #1, MissionNumber + 1
Close #1
tmrLiftoff.Enabled = True
End If
If Countdown = 1 Then
lblEvents.Caption = "T-minus 1 second and counting"
End If
End Sub

Private Sub tmrLiftoff_Timer()
Set objPic1 = LoadPicture(App.Path & "\ShuttleInFlight1.jpg")
Image1.Picture = objPic1
Set objPic1 = LoadPicture(App.Path & "\ShuttleInFlight1.jpg")
Image2.Picture = objPic1
Set objPic1 = LoadPicture(App.Path & "\ShuttleInFlight1.jpg")
Image3.Picture = objPic1
End Sub

Private Sub txtCommander_Change()
Commander = txtCommander.Text
If txtCommander.Text = "" Then
cmdLiftoff.Enabled = False
End If
End Sub

Private Sub txtCoPilot_Change()
CoPilot = txtCoPilot.Text
If txtCoPilot.Text = "" Then
cmdLiftoff.Enabled = False
End If

End Sub

Private Sub txtCountdown_Change()
If Val(txtCountdown.Text) > 0 And cboOrbiters.Text <> "" And MisPurpose <> "" Then
cmdLiftoff.Enabled = True
End If
If Val(txtCountdown.Text) < 1 Or txtCountdown.Text = "" Or cboOrbiters.Text = "" Or MisPurpose = "" Then
cmdLiftoff.Enabled = False
End If
If cboOrbiters.Text = "" Or MisPurpose = "" Then
MsgBox "Choose a mission purpose,crew,and an orbiter first.", 16
txtCountdown.Text = ""
End If
Countdown = txtCountdown.Text
End Sub

Private Sub txtMisSpecial_Change()
MisSpecial = txtMisSpecial.Text
If txtMisSpecial.Text = "" Then
cmdLiftoff.Enabled = False
End If
End Sub

Private Sub txtPayloadSpecial_Change()
PayloadSpecial = txtPayloadSpecial.Text
If txtPayloadSpecial.Text = "" Then
cmdLiftoff.Enabled = False
End If
End Sub

Private Sub txtPilot_Change()
Pilot = txtPilot.Text
If txtPilot.Text = "" Then
cmdLiftoff.Enabled = False
End If
End Sub
