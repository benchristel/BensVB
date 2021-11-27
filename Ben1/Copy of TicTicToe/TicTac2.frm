VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5715
   ClientLeft      =   4950
   ClientTop       =   3165
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5715
   ScaleWidth      =   7065
   Begin VB.Label Label10 
      Caption         =   "Click on a Square"
      Height          =   735
      Left            =   840
      TabIndex        =   9
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3480
      TabIndex        =   8
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   2160
      TabIndex        =   7
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   840
      TabIndex        =   6
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   2160
      TabIndex        =   4
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   840
      TabIndex        =   3
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3480
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub compmove()
If Label5.Caption <> Empty Then GoTo ck3
Label5.Caption = CompMark
Exit Sub
ck3:
If Label3.Caption <> Empty Then GoTo ck1
Label3.Caption = CompMark
Exit Sub
ck1:
If Label1.Caption <> Empty Then GoTo ck7
Label1.Caption = CompMark
Exit Sub
ck7:
If Label7.Caption <> Empty Then GoTo ck9
Label7.Caption = CompMark
Exit Sub
ck9:
If Label9.Caption <> Empty Then GoTo ck2
Label9.Caption = CompMark
Exit Sub
ck2:
If Label2.Caption <> Empty Then GoTo ck4
Label2.Caption = CompMark
Exit Sub
ck4:
If Label4.Caption <> Empty Then GoTo ck6
Label4.Caption = CompMark
Exit Sub
ck6:
If Label6.Caption <> Empty Then GoTo ck8
Label6.Caption = CompMark
Exit Sub
ck8:
If Label8.Caption <> Empty Then Exit Sub
Label8.Caption = CompMark
End Sub

Private Sub Label1_Click()
If Label1.Caption <> Empty Then Exit Sub
Label1.Caption = UserMark
Call compmove
End Sub

Private Sub Label2_Click()
If Label2.Caption <> Empty Then Exit Sub
Label2.Caption = UserMark
Call compmove
End Sub

Private Sub Label3_Click()
If Label3.Caption <> Empty Then Exit Sub
Label3.Caption = UserMark
Call compmove
End Sub

Private Sub Label4_Click()
If Label4.Caption <> Empty Then Exit Sub
Label4.Caption = UserMark
Call compmove
End Sub

Private Sub Label5_Click()
If Label5.Caption <> Empty Then Exit Sub
Label5.Caption = UserMark
Call compmove
End Sub

Private Sub Label6_Click()
If Label6.Caption <> Empty Then Exit Sub
Label6.Caption = UserMark
Call compmove
End Sub

Private Sub Label7_Click()
If Label7.Caption <> Empty Then Exit Sub
Label7.Caption = UserMark
Call compmove
End Sub

Private Sub Label8_Click()
If Label8.Caption <> Empty Then Exit Sub
Label8.Caption = UserMark
Call compmove
End Sub

Private Sub Label9_Click()
If Label9.Caption <> Empty Then Exit Sub
Label9.Caption = UserMark
Call compmove
End Sub

