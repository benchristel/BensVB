VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEquals 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   3255
   End
   Begin VB.CommandButton cmdTimes 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   16
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdOff 
      Caption         =   "off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   15
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   13
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "¸"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   12
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   11
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   10
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblDisplay 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim current, first, answer, number, operation

Private Sub cmd1_Click()
If number = "0" Then
Call Enter1_0
End If
End Sub

Private Sub cmd2_Click()
If current = "0" Then
current = "2"
lblDisplay.Caption = current
current = lblDisplay.Caption
Exit Sub
Else
lblDisplay.Caption = current & "2"
current = lblDisplay.Caption
End If
End Sub

Private Sub cmdEquals_Click()
If operation = "" Then
Exit Sub
End If
If operation = "Plus" Then
answer = first + current
lblDisplay.Caption = answer
Exit Sub
End If
End Sub

Private Sub cmdPlus_Click()
first = lblDisplay.Caption
current = 0
lblDisplay.Caption = "0"
operation = "plus"
End Sub

Private Sub Form_Load()
current = "0"
number = "0"
End Sub

Private Sub Enter1_0()
If current = "0" Then
current = "1"
lblDisplay.Caption = current
current = lblDisplay.Caption
Exit Sub
Else
lblDisplay.Caption = current & "1"
current = lblDisplay.Caption
End If
End Sub
