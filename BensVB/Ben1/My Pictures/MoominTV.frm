VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStop 
      BeginProperty Font 
         Name            =   "French Script MT"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   5520
      Width           =   615
   End
   Begin VB.Label labelHeading 
      Alignment       =   2  'Center
      Caption         =   "Moomin TV"
      BeginProperty Font 
         Name            =   "Blackadder ITC"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   6015
   End
   Begin VB.Image imgTV 
      BorderStyle     =   1  'Fixed Single
      Height          =   4800
      Left            =   1200
      Stretch         =   -1  'True
      Top             =   600
      Width           =   6405
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim FileName, i, j, k, strI, selectedfile, playing
Private Sub cmdPlay_Click()
playing = True
'While playing = True
For j = 1 To 100
    For i = 367 To 390
    FileName = "C:\My Documents\Pictures\train\IMG_"
    strI = Format(i, "0000")
    FileName = FileName & strI & ".jpg"
'    selectedfile = App.Path & "\" & FileName
    selectedfile = FileName
    imgTV.Picture = LoadPicture(selectedfile)
    DoEvents
    Next i
    For k = 1 To 3
    For i = 354 To 364
    FileName = "C:\My Documents\Pictures\train\IMG_"
    strI = Format(i, "0000")
    FileName = FileName & strI & ".jpg"
    selectedfile = App.Path & "\" & FileName
    selectedfile = FileName
    imgTV.Picture = LoadPicture(selectedfile)
    DoEvents
    Next i
    Next k
Next j
'Loop
End Sub

Private Sub cmdStop_Click()
playing = False
End Sub
