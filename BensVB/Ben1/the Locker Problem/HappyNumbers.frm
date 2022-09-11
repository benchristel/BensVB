VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   11250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14865
   LinkTopic       =   "Form2"
   ScaleHeight     =   11250
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtEnd 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "100"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtStart 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblOutput 
      Height          =   11055
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   13215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim inputnum As Integer, outputnum As Integer
Dim prevoutput() As Integer, cycles As Integer
Dim startnum As Integer, endnum As Integer
Dim inputstring As String
Private Sub cmdGo_Click()
Dim i, x
lblOutput.Caption = ""
For i = startnum To endnum
    outputnum = 0
    inputnum = i
    Do Until outputnum = 1
        outputnum = 0
        cycles = cycles + 1
        ReDim Preserve prevoutput(1 To cycles)
        'add squares of digits
        inputstring = inputnum
        For x = 1 To Len(inputstring)
            outputnum = outputnum + (Mid(inputnum, x, 1)) ^ 2
        Next x
        'check for repeats
        For x = 1 To cycles
            If prevoutput(x) = outputnum Then Exit Do
        Next x
        prevoutput(cycles) = outputnum
        'prepare for next cycle
        inputnum = outputnum
    Loop
    If outputnum <> 1 Then GoTo skipprintout
    lblOutput.Caption = lblOutput.Caption & "  " & i
skipprintout:
    cycles = 0
    ReDim prevoutput(1 To 1)
Next i
End Sub

Private Sub Form_Load()
cycles = 0
startnum = Int(Val(txtStart.Text))
endnum = Int(Val(txtEnd.Text))
End Sub

Private Sub txtEnd_Change()
endnum = Int(Val(txtEnd.Text))
End Sub

Private Sub txtStart_Change()
startnum = Int(Val(txtStart.Text))
End Sub
