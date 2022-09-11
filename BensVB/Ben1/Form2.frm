VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12540
   LinkTopic       =   "Form2"
   ScaleHeight     =   5400
   ScaleWidth      =   12540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      Height          =   4575
      Left            =   4860
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   180
      Width           =   7095
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   5040
      Width           =   1395
   End
   Begin VB.TextBox txtInput 
      Height          =   4935
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   4695
   End
   Begin VB.Label lblOutput 
      Caption         =   "Click Here"
      BeginProperty Font 
         Name            =   "Beanie"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4800
      TabIndex        =   1
      Top             =   4860
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WordInput(), Output
Dim WordData(), Words, RandLink()

Private Sub cmdSubmit_Click()
Dim ReadPos As Integer 'the position of the reading head
Dim wordlength, WordFound As Boolean
txtInput.Text = txtInput.Text & " " 'add space to the end to make sure the whole input will be processed.
Do Until Trim(txtInput.Text) = ""
Do Until WordFound = True
ReadPos = ReadPos + 1
If ReadPos > Len(txtInput.Text) Then Exit Sub
If Mid(txtInput.Text, ReadPos, 1) = " " Then WordFound = True
Loop
If Trim(Left(txtInput.Text, ReadPos)) <> "" Then
Words = Words + 1
ReDim Preserve WordData(1 To Words)
WordData(Words) = UCase(Trim(Left(txtInput.Text, ReadPos)))
End If
txtInput.Text = Right(txtInput.Text, Len(txtInput.Text) - ReadPos)
ReadPos = 0
WordFound = False
Loop
End Sub

Private Sub lblOutput_Click()
Dim i 'each variable in the array contains an integer which refers to a word in the database.  This
                    'allows the program to generate non-consecutive random numbers.
Dim LastWord, SecondLast, ThirdLast, OutputLength 'in words
Dim RandWords 'number of word possibilities
Dim Cycles, OutputWord
Dim PlaceHolder, Place2 'the placebo
Randomize
'On Error GoTo 1:
Output = WordData(1) & " " & WordData(2)
OutputLength = OutputLength + 1
LastWord = WordData(2)
SecondLast = WordData(1)
Do Until Cycles = 500
For i = 2 To Words
If WordData(i) = LastWord And WordData(i - 1) = SecondLast And i < Words Then
RandWords = RandWords + 1
ReDim Preserve RandLink(1 To RandWords)
RandLink(RandWords) = i + 1
End If
Next i
If RandWords > 0 Then
OutputWord = WordData(RandLink(Int(Rnd * RandWords + 1)))
'Output = WordData(RandLink(Int(Rnd * RandWords + 1)))
PlaceHolder = LastWord
LastWord = OutputWord
SecondLast = PlaceHolder
Output = Output & " " & OutputWord
End If
Cycles = Cycles + 1
'ReDim RandLink(1 To 1)
RandWords = 0
Loop
1:
txtOutput.Text = Output
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
'If KeyAscii = 32 And Trim(txtInput.Text) <> "" Then 'SPACE key has been pressed.
'Words = Words + 1
'ReDim Preserve WordData(1 To Words)
'WordData(Words) = Trim(txtInput.Text)
'txtInput.Text = ""
'End If
End Sub

