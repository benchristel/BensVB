VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LightningFlash"
   ClientHeight    =   6225
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAccent 
      Caption         =   "Ñ"
      Height          =   375
      Index           =   11
      Left            =   4680
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAccent 
      Caption         =   "ñ"
      Height          =   375
      Index           =   10
      Left            =   1920
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAccent 
      Caption         =   "Ú"
      Height          =   375
      Index           =   9
      Left            =   4320
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAccent 
      Caption         =   "Ó"
      Height          =   375
      Index           =   8
      Left            =   3960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAccent 
      Caption         =   "Í"
      Height          =   375
      Index           =   7
      Left            =   3600
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAccent 
      Caption         =   "É"
      Height          =   375
      Index           =   6
      Left            =   3240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAccent 
      Caption         =   "Á"
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAccent 
      Caption         =   "ú"
      Height          =   375
      Index           =   4
      Left            =   1560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAccent 
      Caption         =   "ó"
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAccent 
      Caption         =   "í"
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAccent 
      Caption         =   "é"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmdAccent 
      Caption         =   "á"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox txtOutputAnswer 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   5400
      Width           =   4935
   End
   Begin VB.TextBox txtOutputGloss 
      BackColor       =   &H00FFFFFF&
      Height          =   2205
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3120
      Width           =   4935
   End
   Begin VB.TextBox txtOutputWord 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   4935
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3615
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Start >>>"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   4935
   End
   Begin VB.Label lblTitle 
      Caption         =   "Now Drilling: None"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblPrompt 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   4935
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuDefault 
         Caption         =   "Set As &Default"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuStack 
      Caption         =   "&Stack"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Card"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Card"
      End
      Begin VB.Menu mnuShuffle 
         Caption         =   "Re-&shuffle"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuInvert 
         Caption         =   "&Flip Deck"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuPrefs 
      Caption         =   "&Preferences"
      Begin VB.Menu mnuTogglePrompts 
         Caption         =   "Editing &Prompts"
         Checked         =   -1  'True
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuShowExplanations 
         Caption         =   "&Show Explanations"
      End
      Begin VB.Menu mnuShowCards 
         Caption         =   "Show &Cards..."
         Begin VB.Menu mnuCardsX 
            Caption         =   "Once"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuCardsX 
            Caption         =   "Twice"
            Index           =   2
         End
         Begin VB.Menu mnuCardsX 
            Caption         =   "Three Times"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Card() As CardData, CardCount As Integer, CurrentCard As Integer, StackName As String
Dim TypeFocus As Integer ' which control has focus
Dim PromptsOn As Boolean 'whether prompts are displayed in textboxes during editing
Dim Inverted As Boolean, ShowExplanations As Boolean, CycleTarget As Integer 'number of times to cycle cards
Dim default As String 'default file to open on startup

Private Sub cmdAccent_Click(Index As Integer)
Select Case TypeFocus
Case Is = 1 'entry box
    txtEntry.Text = txtEntry.Text & cmdAccent(Index).Caption
    txtEntry.SetFocus
    txtEntry.SelStart = Len(txtEntry.Text)
Case Is = 2 'question box
    txtOutputWord.Text = txtOutputWord.Text & cmdAccent(Index).Caption
    txtOutputWord.SetFocus
    txtOutputWord.SelStart = Len(txtOutputWord.Text)
Case Is = 3 'gloss box
    txtOutputGloss.Text = txtOutputGloss.Text & cmdAccent(Index).Caption
    txtOutputGloss.SetFocus
    txtOutputGloss.SelStart = Len(txtOutputGloss.Text)
Case Is = 4 'answer box
    txtOutputAnswer.Text = txtOutputAnswer.Text & cmdAccent(Index).Caption
    txtOutputAnswer.SetFocus
    txtOutputAnswer.SelStart = Len(txtOutputAnswer.Text)
End Select
End Sub

Private Sub cmdEnter_Click()
Call Enter
End Sub

Private Sub cmdNext_Click()
If txtOutputWord.Enabled = True Then
With Card(CurrentCard)
    .Answer = txtOutputAnswer.Text
    .Correct = False
    .Explanation = txtOutputGloss.Text
    .Prompt = txtOutputWord.Text
End With
End If
Call nextcard
End Sub

Private Sub Form_Load()
Dim temp
ReDim Card(1 To 1)
StackName = ""
CurrentCard = 1
CardCount = 1
Open App.Path & "\Preferences.dat" For Input As #1
Line Input #1, temp
default = temp
Line Input #1, temp
ShowExplanations = temp
Line Input #1, temp
PromptsOn = temp
Line Input #1, temp
CycleTarget = temp
Close #1
If default = "" Then
If PromptsOn = True Then
txtOutputGloss.Text = "Explanation of answer"
txtOutputWord.Text = "Prompt to display"
txtOutputAnswer.Text = "Short answer to prompt"
End If
txtEntry.Enabled = False
cmdEnter.Enabled = False
Else
txtOutputGloss.Enabled = False
txtOutputWord.Enabled = False
txtOutputAnswer.Enabled = False
cmdNext.Caption = "Next >>>"
If default <> "" Then
    filename = default
    On Error GoTo filenotfound
    Open App.Path & "\" & filename & ".dat" For Input As #1
    Line Input #1, temp
    StackName = temp
    Line Input #1, temp
    CardCount = temp
    ReDim Card(1 To CardCount)
    For i = 1 To CardCount
    Line Input #1, temp
    Card(i).Prompt = temp
    Line Input #1, temp
    Card(i).Answer = temp
    Line Input #1, temp
    Card(i).Explanation = temp
    Card(i).Correct = 0
    Next i
    Close #1
    Call Shuffle
lblTitle.Caption = "Now Drilling: " & UCase(StackName)
lblPrompt.Caption = CurrentCard & "/" & CardCount & ": " & Card(CurrentCard).Prompt
End If
End If
For i = 1 To 3
mnuCardsX(i).Checked = False
Next i
mnuCardsX(CycleTarget).Checked = True
If PromptsOn = False Then mnuTogglePrompts.Checked = False
If ShowExplanations = True Then mnuShowExplanations.Checked = True
filenotfound:
End Sub

Private Sub Form_Unload(Cancel As Integer)
If StackName <> "" Then
Select Case MsgBox("Do you want to save changes made to '" & StackName & "'?", vbYesNoCancel, "Exit LightningFlash")
Case Is = vbYes
Call Save
Case Is = vbCancel
Cancel = 1
End Select
Else
Select Case MsgBox("Do you want to save changes made to this file?", vbYesNoCancel, "Exit LightningFlash")
Case Is = vbYes
Call SaveAs
Case Is = vbCancel
Cancel = 1
End Select
End If
Open App.Path & "\Preferences.dat" For Output As #1
Print #1, default
Print #1, ShowExplanations
Print #1, PromptsOn
Print #1, CycleTarget
Close #1
End Sub

Private Sub mnuAdd_Click()
If txtOutputWord.Enabled = True Then
With Card(CurrentCard)
    .Answer = txtOutputAnswer.Text
    .Correct = False
    .Explanation = txtOutputGloss.Text
    .Prompt = txtOutputWord.Text
End With
End If
CardCount = CardCount + 1
ReDim Preserve Card(1 To CardCount)
CurrentCard = CardCount
lblPrompt.Caption = ""
Select Case PromptsOn
Case Is = True
txtOutputGloss.Text = "Explanation of answer"
txtOutputWord.Text = "Prompt to display"
txtOutputAnswer.Text = "Short answer to prompt"
Case Is = False
txtOutputGloss.Text = ""
txtOutputWord.Text = ""
txtOutputAnswer.Text = ""
End Select
txtOutputGloss.Enabled = True
txtOutputAnswer.Enabled = True
txtOutputWord.Enabled = True
cmdEnter.Enabled = False
cmdEnter.Caption = "Enter"
cmdNext.Caption = "Start >>>"
txtOutputWord.SetFocus
txtEntry.Enabled = False
End Sub

Private Sub mnuCardsX_Click(Index As Integer)
Dim i
For i = 1 To 3
mnuCardsX(i).Checked = False
Next i
mnuCardsX(Index).Checked = True
CycleTarget = Index
End Sub

Private Sub mnuDefault_Click()
Select Case StackName
Case Is = ""
If MsgBox("Do you want a new file to open each time you start the program?", vbYesNo, "Set New as Default") = vbYes Then
default = ""
End If
Case Else
If MsgBox("Do you want to set this stack to open each time you start the program?", vbYesNo, "Set as Default") = vbYes Then
default = StackName
End If
End Select


End Sub

Private Sub mnuDelete_Click()
Dim temp As CardData, i As Integer
If CardCount = 1 Then
MsgBox "Can't remove the only card in a stack.", vbCritical, "Error!"
Exit Sub
End If
If MsgBox("Are you sure you want to remove this card?", vbYesNo, "Remove Card") = vbYes Then
For i = CurrentCard To CardCount
If i < CardCount Then Card(i) = Card(i + 1)
Next i
CardCount = CardCount - 1
ReDim Preserve Card(1 To CardCount)
End If
If CurrentCard > CardCount Then CurrentCard = CardCount
lblPrompt.Caption = CurrentCard & "/" & CardCount & ": " & Card(CurrentCard).Prompt
txtEntry.Text = ""
txtOutputGloss.Text = ""
txtOutputAnswer.Text = ""
txtOutputWord.Text = ""
cmdEnter.Enabled = True
    If ShowExplanations = True Then txtOutputGloss.Text = Card(CurrentCard).Explanation
cmdNext.Caption = "Next >>>"
txtEntry.SetFocus
End Sub

Private Sub mnuDrillAll_Click()

End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuInvert_Click()
If MsgBox("Flipping the stack will cause you to review all cards." & vbLf & "Do you want to continue?", vbYesNo, "Flip Cards") = vbYes Then
Select Case Inverted
Case Is = True
    Inverted = False
    mnuInvert.Checked = False
Case Is = False
    Inverted = True
    mnuInvert.Checked = True
End Select
Call Shuffle
End If
End Sub

Private Sub mnuNew_Click()
If StackName <> "" Then
If MsgBox("Do you want to save changes made to '" & StackName & "'?", vbYesNo, "Create New File") = vbYes Then
Call Save
End If
Else
If MsgBox("Do you want to save changes made to this file?", vbYesNoCancel, "Create New File") = vbYes Then
Call SaveAs
End If
End If
ReDim Card(1 To 1)
CardCount = 1
With Card(1)
    .Answer = ""
    .Correct = 0
    .Explanation = ""
    .Prompt = ""
End With
Select Case PromptsOn
Case Is = True
txtOutputGloss.Text = "Explanation of answer"
txtOutputWord.Text = "Prompt to display"
txtOutputAnswer.Text = "Short answer to prompt"
Case Is = False
txtEntry.Text = ""
txtOutputGloss.Text = ""
txtOutputWord.Text = ""
txtOutputAnswer.Text = ""
End Select
lblPrompt.Caption = ""
lblTitle.Caption = "Now Drilling: None"
StackName = ""
CurrentCard = 1
CardCount = 1
cmdEnter.Enabled = False
txtOutputWord.Enabled = True
txtOutputGloss.Enabled = True
txtOutputAnswer.Enabled = True
txtOutputWord.SetFocus
txtEntry.Enabled = False
cmdEnter.Caption = "Enter"
cmdNext.Caption = "Start >>>"
End Sub

Private Sub mnuOpen_Click()
Dim filename As String, temp, i
If StackName <> "" Then
If MsgBox("Do you want to save changes made to '" & StackName & "'?", vbYesNo, "Exit LightningFlash") = vbYes Then
Call Save
End If
End If
filename = LCase(InputBox("Type the name of the stack you wish to open.", "Open File"))
On Error GoTo filenotfound
Open App.Path & "\" & filename & ".dat" For Input As #1
Line Input #1, temp
StackName = temp
Line Input #1, temp
CardCount = temp
ReDim Card(1 To CardCount)
For i = 1 To CardCount
Line Input #1, temp
Card(i).Prompt = temp
Line Input #1, temp
Card(i).Answer = temp
Line Input #1, temp
Card(i).Explanation = temp
Card(i).Correct = 0
Next i
Close #1
Call Shuffle
CurrentCard = 1
lblTitle.Caption = "Now Drilling: " & UCase(StackName)
Select Case Inverted
Case Is = False ' show prompt, expect answer
    lblPrompt.Caption = CurrentCard & "/" & CardCount & ": " & Card(CurrentCard).Prompt
Case Is = True 'show answer, expect prompt
    lblPrompt.Caption = CurrentCard & "/" & CardCount & ": " & Card(CurrentCard).Answer
End Select
txtEntry.Text = ""
txtEntry.Enabled = True
txtOutputGloss.Text = ""
txtOutputAnswer.Text = ""
txtOutputWord.Text = ""
txtOutputGloss.Enabled = False
txtOutputAnswer.Enabled = False
txtOutputWord.Enabled = False
    If ShowExplanations = True Then txtOutputGloss.Text = Card(CurrentCard).Explanation
cmdEnter.Enabled = True
cmdEnter.Caption = "Show Answer"
cmdNext.Caption = "Next >>>"
txtEntry.SetFocus
Exit Sub
filenotfound:
MsgBox "Could not find the specified file.", vbCritical, "Error!"
End Sub

Private Sub mnuSave_Click()
Dim i
If cmdEnter.Enabled = False Then
With Card(CurrentCard)
    .Answer = txtOutputAnswer.Text
    .Correct = False
    .Explanation = txtOutputGloss.Text
    .Prompt = txtOutputWord.Text
End With
End If
If StackName = "" Then
Call SaveAs
Exit Sub
End If
Call Save
End Sub

Private Sub mnuSaveAs_Click()
If cmdEnter.Enabled = False Then
With Card(CurrentCard)
    .Answer = txtOutputAnswer.Text
    .Correct = False
    .Explanation = txtOutputGloss.Text
    .Prompt = txtOutputWord.Text
End With
End If
Call SaveAs
End Sub

Private Sub mnuShowExplanations_Click()
Select Case ShowExplanations
Case Is = True
    ShowExplanations = False
    mnuShowExplanations.Checked = False
Case Is = False
    ShowExplanations = True
    mnuShowExplanations.Checked = True
End Select
End Sub

Private Sub mnuShuffle_Click()
If MsgBox("Shuffling will cause you to review all cards in this stack.  Do you want to continue?", vbYesNo, "Shuffle") = vbYes Then
Call Shuffle
End If
End Sub

Public Sub Shuffle()
Dim i, temp As CardData, card1 As Integer, card2 As Integer
For i = 1 To CardCount
Card(i).Correct = 0
Next i
Randomize
For i = 1 To CardCount * 2
card1 = Int(Rnd * CardCount + 1)
card2 = Int(Rnd * CardCount + 1)
temp = Card(card1)
Card(card1) = Card(card2)
Card(card2) = temp
Next i
CurrentCard = 1
Select Case Inverted
Case Is = False ' show prompt, expect answer
    lblPrompt.Caption = CurrentCard & "/" & CardCount & ": " & Card(CurrentCard).Prompt
Case Is = True 'show answer, expect prompt
    lblPrompt.Caption = CurrentCard & "/" & CardCount & ": " & Card(CurrentCard).Answer
End Select
txtEntry.Text = ""
txtOutputGloss.Text = ""
txtOutputAnswer.Text = ""
txtOutputWord.Text = ""
    If ShowExplanations = True Then txtOutputGloss.Text = Card(CurrentCard).Explanation
cmdEnter.Enabled = True
End Sub


Public Sub SaveAs()
Dim i
For i = 1 To CardCount
Card(i).Explanation = Replace(Card(i).Explanation, vbCr, " ")
Card(i).Explanation = Replace(Card(i).Explanation, vbLf, " ")
Next i
StackName = LCase(InputBox("Enter a name for this flashcard stack:", "Save As"))
If Len(StackName) = 0 Or Len(StackName) > 30 Then
MsgBox "Please enter a unique filename between 1 and 30 characters long.", vbCritical, "Error!"
StackName = ""
Exit Sub
End If
Open App.Path & "\" & StackName & ".dat" For Output As #1
Print #1, StackName
Print #1, CardCount
For i = 1 To CardCount
Print #1, Card(i).Prompt
Print #1, Card(i).Answer
Print #1, Card(i).Explanation
Next i
Close #1
End Sub

Private Sub mnuTogglePrompts_Click()
Select Case PromptsOn
Case Is = True
    PromptsOn = False
    mnuTogglePrompts.Checked = False
Case Is = False
    PromptsOn = True
    mnuTogglePrompts.Checked = True
End Select
End Sub

Private Sub txtEntry_Change()
If Trim(txtEntry.Text) = "" And cmdEnter.Enabled = True Then
cmdEnter.Caption = "Show Answer"
Else
cmdEnter.Caption = "Enter"
End If
End Sub

Private Sub txtEntry_GotFocus()
TypeFocus = 1
End Sub

Private Sub txtEntry_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And cmdEnter.Enabled = True Then 'enter was pressed
Call Enter
End If
If KeyCode = 34 Then 'page down
Call nextcard
End If
End Sub

Private Sub txtOutputAnswer_GotFocus()
TypeFocus = 4
End Sub

Private Sub txtOutputGloss_GotFocus()
TypeFocus = 3
End Sub

Private Sub txtOutputWord_GotFocus()
TypeFocus = 2
End Sub

Public Sub nextcard()
Dim i
cmdNext.Caption = "Next >>>"
If Card(CurrentCard).Prompt = "" Then
With Card(CurrentCard)
    .Answer = txtOutputAnswer.Text
    .Correct = False
    .Explanation = txtOutputGloss.Text
    .Prompt = txtOutputWord.Text
End With
End If
For i = CurrentCard + 1 To CardCount
If Card(i).Correct < CycleTarget Then
CurrentCard = i
GoTo showcard
End If
Next i
For i = 1 To CurrentCard
If Card(i).Correct < CycleTarget Then
CurrentCard = i
GoTo showcard
End If
Next i
MsgBox "You have correctly answered all the cards in this stack." & vbLf & "You will continue drilling this stack unless you load a new one", vbOKOnly, "Stack Complete"
For i = 1 To CardCount
Card(i).Correct = 0
Next i
CurrentCard = 1
Call Shuffle
showcard:
'====================
Select Case Inverted
Case Is = False ' show prompt, expect answer
    lblPrompt.Caption = CurrentCard & "/" & CardCount & ": " & Card(CurrentCard).Prompt
Case Is = True 'show answer, expect prompt
    lblPrompt.Caption = CurrentCard & "/" & CardCount & ": " & Card(CurrentCard).Answer
End Select
    txtEntry.Text = ""
    txtEntry.Enabled = True
    txtOutputGloss.Text = ""
    txtOutputAnswer.Text = ""
    txtOutputWord.Text = ""
    txtOutputGloss.Enabled = False
    txtOutputAnswer.Enabled = False
    txtOutputWord.Enabled = False
    cmdEnter.Enabled = True
    cmdEnter.Caption = "Show Answer"
    If ShowExplanations = True Then txtOutputGloss.Text = Card(CurrentCard).Explanation
    txtEntry.SetFocus
End Sub

Public Sub Save()
Dim i
If StackName = "" Then Call SaveAs
For i = 1 To CardCount
Card(i).Explanation = Replace(Card(i).Explanation, vbCr, " ")
Card(i).Explanation = Replace(Card(i).Explanation, vbLf, " ")
Next i
Open App.Path & "\" & StackName & ".dat" For Output As #1
Print #1, StackName
Print #1, CardCount
For i = 1 To CardCount
Print #1, Card(i).Prompt
Print #1, Card(i).Answer
Print #1, Card(i).Explanation
Next i
Close #1
End Sub

Public Sub Enter()
If (LCase(Trim(txtEntry.Text)) = LCase(Trim(Card(CurrentCard).Answer)) And Inverted = False) Or (LCase(Trim(txtEntry.Text)) = LCase(Trim(Card(CurrentCard).Prompt)) And Inverted = True) Then
Card(CurrentCard).Correct = Card(CurrentCard).Correct + 1
lblPrompt.Caption = "Correct!"
ElseIf Trim(txtEntry.Text) = "" Then
If Inverted = False Then
lblPrompt.Caption = Card(CurrentCard).Answer
Else
lblPrompt.Caption = Card(CurrentCard).Prompt
End If
Else
If Inverted = False Then
lblPrompt.Caption = "Sorry, the correct answer is '" & Card(CurrentCard).Answer & "'"
Else
lblPrompt.Caption = "Sorry, the correct answer is '" & Card(CurrentCard).Prompt & "'"
End If
End If
cmdEnter.Enabled = False
txtOutputWord.Text = Card(CurrentCard).Prompt
txtOutputGloss.Text = Card(CurrentCard).Explanation
txtOutputAnswer.Text = Card(CurrentCard).Answer
txtOutputGloss.Enabled = True
txtOutputAnswer.Enabled = True
txtOutputWord.Enabled = True

End Sub
