VERSION 5.00
Begin VB.Form frmEditor 
   Caption         =   "Form1"
   ClientHeight    =   11730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13290
   LinkTopic       =   "Form1"
   ScaleHeight     =   11730
   ScaleWidth      =   13290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTimeSig 
      Caption         =   "4/4 Mode"
      Height          =   375
      Left            =   6240
      TabIndex        =   90
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtGoto 
      Height          =   375
      Left            =   840
      TabIndex        =   89
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export"
      Height          =   375
      Left            =   6600
      TabIndex        =   87
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open File:"
      Height          =   375
      Left            =   6600
      TabIndex        =   86
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox txtOpenFile 
      Height          =   375
      Left            =   7800
      TabIndex        =   85
      Text            =   "DefaultProject"
      Top             =   5760
      Width           =   3855
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save As:"
      Height          =   375
      Left            =   6600
      TabIndex        =   84
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   ">>"
      Height          =   375
      Left            =   6840
      TabIndex        =   83
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<"
      Height          =   375
      Left            =   5880
      TabIndex        =   82
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtFilename 
      Height          =   375
      Left            =   7800
      TabIndex        =   81
      Text            =   "DefaultProject"
      Top             =   5280
      Width           =   3855
   End
   Begin VB.TextBox lblDuration 
      Height          =   375
      Left            =   2040
      TabIndex        =   80
      Text            =   "0"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblMeasure 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   88
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   79
      Left            =   4440
      TabIndex        =   79
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   78
      Left            =   3840
      TabIndex        =   78
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   77
      Left            =   3240
      TabIndex        =   77
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   76
      Left            =   2640
      TabIndex        =   76
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   75
      Left            =   2040
      TabIndex        =   75
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   74
      Left            =   4440
      TabIndex        =   74
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   73
      Left            =   3840
      TabIndex        =   73
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   72
      Left            =   3240
      TabIndex        =   72
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   71
      Left            =   2640
      TabIndex        =   71
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   70
      Left            =   2040
      TabIndex        =   70
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   69
      Left            =   4440
      TabIndex        =   69
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   68
      Left            =   3840
      TabIndex        =   68
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   67
      Left            =   3240
      TabIndex        =   67
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   66
      Left            =   2640
      TabIndex        =   66
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   65
      Left            =   2040
      TabIndex        =   65
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   64
      Left            =   4440
      TabIndex        =   64
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   63
      Left            =   3840
      TabIndex        =   63
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   62
      Left            =   3240
      TabIndex        =   62
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   61
      Left            =   2640
      TabIndex        =   61
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   60
      Left            =   2040
      TabIndex        =   60
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   59
      Left            =   4440
      TabIndex        =   59
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   58
      Left            =   3840
      TabIndex        =   58
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   57
      Left            =   3240
      TabIndex        =   57
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   56
      Left            =   2640
      TabIndex        =   56
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   55
      Left            =   2040
      TabIndex        =   55
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   54
      Left            =   4440
      TabIndex        =   54
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   53
      Left            =   3840
      TabIndex        =   53
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   52
      Left            =   3240
      TabIndex        =   52
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   51
      Left            =   2640
      TabIndex        =   51
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   50
      Left            =   2040
      TabIndex        =   50
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   49
      Left            =   4440
      TabIndex        =   49
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   48
      Left            =   3840
      TabIndex        =   48
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   47
      Left            =   3240
      TabIndex        =   47
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   46
      Left            =   2640
      TabIndex        =   46
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   45
      Left            =   2040
      TabIndex        =   45
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   44
      Left            =   4440
      TabIndex        =   44
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   43
      Left            =   3840
      TabIndex        =   43
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   42
      Left            =   3240
      TabIndex        =   42
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   41
      Left            =   2640
      TabIndex        =   41
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   40
      Left            =   2040
      TabIndex        =   40
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   39
      Left            =   4440
      TabIndex        =   39
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   38
      Left            =   3840
      TabIndex        =   38
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   37
      Left            =   3240
      TabIndex        =   37
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   36
      Left            =   2640
      TabIndex        =   36
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   35
      Left            =   2040
      TabIndex        =   35
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   34
      Left            =   4440
      TabIndex        =   34
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   33
      Left            =   3840
      TabIndex        =   33
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   32
      Left            =   3240
      TabIndex        =   32
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   31
      Left            =   2640
      TabIndex        =   31
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   30
      Left            =   2040
      TabIndex        =   30
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   29
      Left            =   4440
      TabIndex        =   29
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   28
      Left            =   3840
      TabIndex        =   28
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   27
      Left            =   3240
      TabIndex        =   27
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   26
      Left            =   2640
      TabIndex        =   26
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   25
      Left            =   2040
      TabIndex        =   25
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   24
      Left            =   4440
      TabIndex        =   24
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   23
      Left            =   3840
      TabIndex        =   23
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   22
      Left            =   3240
      TabIndex        =   22
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   21
      Left            =   2640
      TabIndex        =   21
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   20
      Left            =   2040
      TabIndex        =   20
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   19
      Left            =   4440
      TabIndex        =   19
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   18
      Left            =   3840
      TabIndex        =   18
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   17
      Left            =   3240
      TabIndex        =   17
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   16
      Left            =   2640
      TabIndex        =   16
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   15
      Left            =   2040
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   14
      Left            =   4440
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   13
      Left            =   3840
      TabIndex        =   13
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   12
      Left            =   3240
      TabIndex        =   12
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   11
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   10
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   9
      Left            =   4440
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   8
      Left            =   3840
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   7
      Left            =   3240
      TabIndex        =   7
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   6
      Left            =   2640
      TabIndex        =   6
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   4
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   3
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Timecode() As Timecode, TimecodeMax As Integer
Dim TimeScroll As Integer
Dim SetDuration As Double
Dim timesig As Integer

Private Sub cmdBack_Click()
TimeScroll = TimeScroll - 4 * timesig
If TimeScroll < 1 Then TimeScroll = 1
UpdateDisplay
End Sub

Private Sub cmdExport_Click()
Dim makefilename As String, musicfile As String, bpm As Double, gamefilecount As Integer, temp, songfile() As String, levelname As String
Dim i As Integer, k As Integer
Dim Note() As Note, notecount As Integer
For i = 1 To TimecodeMax
For k = 0 To 4
If Timecode(i).Note(k) = True Then
    notecount = notecount + 1
    ReDim Preserve Note(1 To notecount)
    With Note(notecount)
    .Distance = (i - 1) / 4
    .Duration = Timecode(i).Duration(k)
    .XOffset = k
    End With
End If
Next k
Next i
Open App.Path & "\Data.dat" For Input As #1
Line Input #1, temp
gamefilecount = temp
ReDim songfile(1 To gamefilecount + 1)
For i = 1 To gamefilecount
Line Input #1, temp
songfile(i) = temp
Next i
Close #1
songfile(gamefilecount + 1) = InputBox("Type a filename to save the song under.", "Export Level", "DefaultSong")
levelname = InputBox("Type a name for this level; this will appear in the game menu.", "Export Level", "My Song")
musicfile = InputBox("Type the filename of a wave file in the 'Music Files' folder.  Do not include the .wav extension.", "Export Level")
bpm = Val(InputBox("Enter the tempo of the song you chose in beats per minute.", "Export Level"))
If MsgBox("Do you want to create a link to this level on the in-game song selection menu?  If you click yes, a new menu item will be created even if one already exists for this level.", vbYesNo, "Export Level") = vbYes Then
    Open App.Path & "\Data.dat" For Output As #1
    Print #1, gamefilecount + 1
    For i = 1 To gamefilecount + 1
    Print #1, songfile(i)
    Next i
    Close #1
End If

Open App.Path & "\Level Files\" & songfile(gamefilecount + 1) & ".dat" For Output As #1
Print #1, levelname
Print #1, musicfile
Print #1, bpm
Print #1, notecount
For i = 1 To notecount
Print #1, Note(i).Distance
Print #1, Note(i).Duration
Print #1, Note(i).XOffset
Next i
Close #1
End Sub

Private Sub cmdForward_Click()
TimeScroll = TimeScroll + 4 * timesig
UpdateDisplay
End Sub

Private Sub cmdOpen_Click()
Dim i, k, temp
Open App.Path & "\Editor Files\" & txtOpenFile.Text & ".dat" For Input As #1
Line Input #1, temp
TimecodeMax = temp
ReDim Timecode(1 To TimecodeMax)
For i = 1 To TimecodeMax
For k = 0 To 4
Line Input #1, temp
Timecode(i).Note(k) = temp
Line Input #1, temp
Timecode(i).Duration(k) = temp
Next k
Next i
Close #1
UpdateDisplay
End Sub

Private Sub cmdSaveAs_Click()
Dim i, k
Open App.Path & "\Editor Files\" & txtFilename.Text & ".dat" For Output As #1
Print #1, TimecodeMax
For i = 1 To TimecodeMax
For k = 0 To 4
Print #1, Timecode(i).Note(k)
Print #1, Timecode(i).Duration(k)
Next k
Next i
Close #1
End Sub

Private Sub cmdTimeSig_Click()
Select Case timesig
Case Is = 4
timesig = 3
cmdTimeSig.Caption = "3/4 Mode"
Case Else
timesig = 3
cmdTimeSig.Caption = "4/4 Mode"
End Select
End Sub

Private Sub Form_Load()
ReDim Timecode(1 To 1)
TimeScroll = 1
timesig = 4
End Sub

Private Sub lblDuration_Change()
SetDuration = Int(Val(lblDuration.Text) * 4) / 4
End Sub

Private Sub lblNote_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
    Case Is = 1 'LMB
        If TimeScroll + Int(Index / 5) > TimecodeMax Then
            TimecodeMax = TimeScroll + Int(Index / 5)
            ReDim Preserve Timecode(1 To TimecodeMax)
        End If
        With Timecode(TimeScroll + Int(Index / 5))
            .Note(Index Mod 5) = True
            .Duration(Index Mod 5) = SetDuration
        End With
    Case Is = 2 'RMB
    On Error GoTo out
        With Timecode(TimeScroll + Int(Index / 5))
            .Note(Index Mod 5) = False
            .Duration(Index Mod 5) = 0
        End With
    End Select
out:
UpdateDisplay
End Sub

Public Sub UpdateDisplay()
Dim i As Integer, k As Integer
On Error GoTo out
For i = TimeScroll To TimeScroll + 15
If i > TimecodeMax Then
    For k = 0 To 4
        lblNote(5 * (i - TimeScroll) + k).BackColor = &H0
    Next k
    GoTo nextrow
End If
For k = 0 To 4
    If Timecode(i).Note(k) = True And lblNote(5 * (i - TimeScroll) + k).BackColor = &H0 Then
        lblNote(5 * (i - TimeScroll) + k).BackColor = RGB(0, 255, 255)
    End If
    If Timecode(i).Note(k) = False And lblNote(5 * (i - TimeScroll) + k).BackColor = RGB(0, 255, 255) Then
        lblNote(5 * (i - TimeScroll) + k).BackColor = &H0
    End If
    lblNote(5 * (i - TimeScroll) + k).Caption = Timecode(i).Duration(k)
Next k
nextrow:
Next i
out:
lblMeasure.Caption = (TimeScroll + 15) / 16
End Sub

Private Sub txtGoto_Change()
If Int(Val(txtGoto.Text) - 1 * 16 + 1) <= 0 Then Exit Sub
TimeScroll = Int(Val(txtGoto.Text - 1) * 16 + 1)
UpdateDisplay
End Sub
