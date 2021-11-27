VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   14550
   ClientLeft      =   150
   ClientTop       =   180
   ClientWidth     =   18735
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   970
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1249
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picScreen 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   14895
      Left            =   0
      ScaleHeight     =   993
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1273
      TabIndex        =   0
      Top             =   0
      Width           =   19095
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   120
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Particle &Options"
      Begin VB.Menu mnuTrails 
         Caption         =   "&Trails"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuGravity 
         Caption         =   "&Gravity"
         Begin VB.Menu mnuQuadratic 
            Caption         =   "&Quadratic"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuLinear 
            Caption         =   "&Linear"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
Dim i
Randomize
loadstars (40)
gravity = 2
End Sub

Private Sub mnuClear_Click()
picScreen.Cls
End Sub

Private Sub mnuExit_Click()
Unload Form1
End
End Sub

Private Sub mnuLinear_Click()
mnuLinear.Checked = True
mnuQuadratic.Checked = False
gravity = 1
End Sub

Private Sub mnuNew_Click()
Dim stars As Integer
stars = Int(Val(InputBox("Enter number of particles to create:", "New Simulation", "40")))
If stars > 0 And stars < 256 Then
loadstars (stars)
Else
MsgBox "Please enter a number between 1 and 255.", , "Error"
End If
End Sub

Private Sub mnuQuadratic_Click()
mnuQuadratic.Checked = True
mnuLinear.Checked = False
gravity = 2
End Sub

Private Sub mnuTrails_Click()
If trails = False Then
mnuTrails.Checked = True
trails = True
Else
mnuTrails.Checked = False
trails = False
End If
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i
For i = 1 To starcount
If Distance((x), (y), Star(i).XCoord, Star(i).YCoord) < 5 Then
picScreen.ToolTipText = i & " " & Star(i).Mass
End If
Next i
End Sub

Private Sub Form_Resize()
picScreen.Width = Form1.Width
picScreen.Height = Form1.Height
End Sub

Private Sub Timer1_Timer()
Dim i, j, tempmagnitude As Double, tempslope As Slope
For i = 1 To starcount
    'If Star(i).XCoord > 1280 Then Star(i).XCoord = 0
    'If Star(i).YCoord > 1000 Then Star(i).YCoord = 0
    'If Star(i).XCoord < 0 Then Star(i).XCoord = 1280
    'If Star(i).YCoord < 0 Then Star(i).YCoord = 1000
    For j = 1 To starcount
        If i = j Then GoTo skip ' stars are not affected by their own gravity
        tempmagnitude = (Star(j).Mass / (Distance(Star(i).XCoord, Star(i).YCoord, Star(j).XCoord, Star(j).YCoord)) ^ gravity) / Star(i).Mass
        tempslope = FindLine(Star(i).XCoord, Star(i).YCoord, Star(j).XCoord, Star(j).YCoord, tempmagnitude)
        Star(i).XVector = Star(i).XVector + tempslope.Run '/ Star(i).Mass ^ 2
        Star(i).YVector = Star(i).YVector + tempslope.Rise '/ Star(i).Mass ^ 2
skip:
    Next j

Next i
For i = 1 To starcount
        Star(i).LastX = Star(i).XCoord
        Star(i).LastY = Star(i).YCoord
        Star(i).XCoord = Star(i).XCoord + Star(i).XVector
        Star(i).YCoord = Star(i).YCoord + Star(i).YVector
Next i
For i = 1 To starcount
    'BitBlt picScreen.hDC, Star(i).LastX, Star(i).LastY, 1, 1, picScreen.hDC, 0, 0, vbBlackness
    'BitBlt picScreen.hDC, Star(i).XCoord, Star(i).YCoord, 1, 1, picScreen.hDC, 0, 0, vbWhiteness
    If trails = False Then picScreen.Circle (Star(i).LastX, Star(i).LastY), 3, vbBlack
Next i
For i = 1 To starcount
    picScreen.Circle (Star(i).XCoord, Star(i).YCoord), 3, Star(i).Color
Next i
    'BitBlt picScreen.hDC, 100, 100, 1, 1, picScreen.hDC, 0, 0, vbWhiteness
Refresh
End Sub

Public Sub loadstars(stars As Integer)
starcount = stars
ReDim Star(1 To starcount)
For i = 1 To starcount
With Star(i)
'    .XCoord = Rnd * Form1.ScaleWidth / 5 - Form1.ScaleWidth / 10 + Form1.ScaleWidth / 2
'    .YCoord = Rnd * Form1.ScaleHeight / 5 - Form1.ScaleHeight / 10 + Form1.ScaleHeight / 2
    .XCoord = Rnd * Form1.ScaleWidth
    .YCoord = Rnd * Form1.ScaleHeight
    .Mass = Int(Rnd * 10000)
    .XVector = (Rnd * 1000 - 500) / Star(i).Mass
    .YVector = (Rnd * 1000 - 500) / Star(i).Mass
    .Color = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
End With
Next i
End Sub
