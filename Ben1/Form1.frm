VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdGoTo 
      Caption         =   "Go To Shelf..."
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtQty 
      BackColor       =   &H80000000&
      Height          =   345
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txtPrice 
      BackColor       =   &H80000000&
      Height          =   345
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox txtProduct 
      BackColor       =   &H80000000&
      Height          =   345
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantity:"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblPrice 
      Alignment       =   1  'Right Justify
      Caption         =   "Price:"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label lblProduct 
      Alignment       =   1  'Right Justify
      Caption         =   "Product:"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
Call Write_Record(ShelfNumber)
If ShelfNumber > 1 Then
ShelfNumber = ShelfNumber - 1
Call Read_Record(ShelfNumber)
Else
Beep
MsgBox "No 'Shelf 0'", 48, "Illegal response"
End If
End Sub

Private Sub cmdGoTo_Click()
Call Write_Record(ShelfNumber)
Dim Result As String, Ans As Integer
Result = InputBox("Go to which shelf?")
Ans = Val(Result)
If Ans >= 1 And Ans <= 100 Then
Call Read_Record(Ans)
Else
MsgBox "You must give a number between 1 and 100.", 48, "Illegal response"
End If
End Sub

Private Sub cmdNext_Click()
Call Write_Record(ShelfNumber)
If ShelfNumber < 100 Then
ShelfNumber = ShelfNumber + 1
Call Read_Record(ShelfNumber)
Else
Beep
MsgBox "Only 100 shelves", 48, "Illegal response"
End If
End Sub

Private Sub Form_Load()
Form1.Visible = False
Form2.Visible = True
Open "C:\Program Files\Microsoft Visual Studio\VB98\BensVB\Ben1\Shelf\Shelf.dat" For Random As #1 Len = Len(Shelf)
Call Read_Record(1)
txtProduct.Text = Shelf.Product
txtPrice.Text = FormatNumber(Shelf.Price, 2)
txtQty.Text = Shelf.Qty
End Sub
Sub Form_Unload(Cancel As Integer)
Close #1
End Sub

Private Sub mnuExit_Click()
Unload Form1
End Sub

Private Sub mnuSave_Click()

End Sub

Private Sub txtQty_Change()
If txtQty.Text = "0" Then Exit Sub
If Val(txtQty.Text) > 0 And Val(txtQty.Text) <= 32000 Then
Exit Sub
Else
MsgBox "can only have 0 to 32000 on a shelf", 48, "Quantity out of range"
End If
End Sub
