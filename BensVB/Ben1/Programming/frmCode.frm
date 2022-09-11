VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDecryptWithoutKey 
      Caption         =   "Decrypt Without Encoding Key"
      Height          =   495
      Left            =   7440
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox txtCodeBase 
      Height          =   315
      Left            =   7440
      TabIndex        =   4
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox txtEncoded 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2640
      Width           =   7350
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "Decrypt"
      Height          =   435
      Left            =   7440
      TabIndex        =   2
      Top             =   540
      Width           =   2655
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encrypt"
      Height          =   435
      Left            =   7440
      TabIndex        =   1
      Top             =   60
      Width           =   2655
   End
   Begin VB.TextBox txtEntered 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   7350
   End
   Begin VB.Label lblStatus 
      Caption         =   "Program Initialized."
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ChrValue(), CodeBase
Private Sub Text1_Change()

End Sub

Private Sub txtCode_Change()

End Sub

Private Sub cmdDecode_Click()
Dim i
On Error GoTo 1
lblStatus.Caption = "Decrypting..."
Rnd (-1)
Randomize (CodeBase)
txtEncoded.Text = ""
ReDim ChrValue(1 To Len(txtEntered.Text))
For i = 1 To Len(txtEntered.Text)
ChrValue(i) = Asc(Mid(txtEntered.Text, i, 1))
ChrValue(i) = ChrValue(i) - Int(Rnd * 10 - 5)
txtEncoded.Text = txtEncoded.Text & Chr(ChrValue(i))
Next i
Refresh
lblStatus.Caption = "Decryption complete."
Exit Sub
1:
lblStatus.Caption = "An error occurred.  Decryption could not be completed."
End Sub

Private Sub cmdDecryptWithoutKey_Click()
Dim i, j
On Error GoTo 1
For j = 1 To CodeBase
lblStatus.Caption = "Decrypting " & j & " of " & CodeBase & "..."
Rnd (-1)
Randomize (j)
txtEncoded.Text = txtEncoded.Text & vbCrLf & j & ": "
ReDim ChrValue(1 To Len(txtEntered.Text))
For i = 1 To Len(txtEntered.Text)
ChrValue(i) = Asc(Mid(txtEntered.Text, i, 1))
ChrValue(i) = ChrValue(i) - Int(Rnd * 10 - 5)
txtEncoded.Text = txtEncoded.Text & Chr(ChrValue(i))
Next i
Refresh
Next j
lblStatus.Caption = "Decryption complete."
Exit Sub
1:
lblStatus.Caption = "An error occurred.  Decryption could not be completed."
End Sub

Private Sub cmdEncode_Click()
Dim i
On Error GoTo 1
lblStatus.Caption = "Encrypting..."
Rnd (-1)
Randomize (CodeBase)
txtEncoded.Text = ""
ReDim ChrValue(1 To Len(txtEntered.Text))
For i = 1 To Len(txtEntered.Text)
ChrValue(i) = Asc(Mid(txtEntered.Text, i, 1))
ChrValue(i) = ChrValue(i) + Int(Rnd * 10 - 5)
txtEncoded.Text = txtEncoded.Text & Chr(ChrValue(i))
Next i
Refresh
lblStatus.Caption = "Encryption complete."
Exit Sub
1:
lblStatus.Caption = "An error occurred.  Encryption could not be completed."
End Sub

Private Sub txtCodeBase_Change()
CodeBase = Int(Val(txtCodeBase.Text))
End Sub
