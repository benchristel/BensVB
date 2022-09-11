Attribute VB_Name = "Module1"
Public Score As Long, RunBonus As Integer, Run As Integer, CumScore As Long, AddScore As Double
Public KeyPressed(1 To 5) As Boolean, MisHit(1 To 5) As Boolean, Strum(1 To 3) As Boolean, StrumTime As Integer
Public Note() As Note, NoteMin As Integer, NoteCount As Integer, NoteMax As Integer
Public NoteDC As Long, NoteMaskDC As Long, HoleDC As Long, HoleMaskDC As Long, HoleLightDC As Long, BackgroundDC As Long, BackgroundImageDC As Long, BackgroundImage2DC As Long, TailDC As Long, HoleLightMaskDC As Long
Public SpinDC(1 To 3) As Long, SpinMaskDC(1 To 3) As Long, SpinState(1 To 5) As Single
Public CenterX As Single, CenterY As Single
Public BeatsPerMin As Double, Notespeed As Double
Public CursorX As Integer, ScrollY As Double
Public SongFileName As String
Public ScreenHeight As Integer, Terminated As Boolean
Public songscroll As Integer, LevelFileName As String
Public Player() As Player, ActivePlayer As Integer, PlayerCount  As Integer
Public songcount As Integer, songname() As String, songfile() As String, ActiveSong As Integer
Public Bloop(1 To 5) As Long
Type Note
    Distance As Double
    XOffset As Integer
    Duration As Double
    Hit As Boolean
    Deleted As Boolean
End Type
Type Player
    Name As String
    MaxScore() As Long
    TotalScore As Long
    VisualOffset As Single
    ResponseOffset As Single
    VisualEffects As Boolean
    SoundEffects As Boolean
End Type
Const KEY_PRESSED As Integer = &H1000
Public Sub UpdateKeys()
Dim i As Integer
'For i = 3 To 200
'If (GetKeyState(i) And KEY_PRESSED) Then MsgBox "you pressed " & i
'Next i
If (GetKeyState(49) And KEY_PRESSED) Then '1 key was pressed
    If KeyPressed(1) = False Then
    Call HitNotes(1, True)
    KeyPressed(1) = True
    End If
Else
    KeyPressed(1) = False
End If
If (GetKeyState(50) And KEY_PRESSED) Then '2 key was pressed
    If KeyPressed(2) = False Then
    Call HitNotes(2, True)
    KeyPressed(2) = True
    End If
Else
    KeyPressed(2) = False
End If
If (GetKeyState(51) And KEY_PRESSED) Then '3 key was pressed
    If KeyPressed(3) = False Then
    Call HitNotes(3, True)
    KeyPressed(3) = True
    End If
Else
    KeyPressed(3) = False
End If
If (GetKeyState(52) And KEY_PRESSED) Then '4 key was pressed
    If KeyPressed(4) = False Then
    Call HitNotes(4, True)
    KeyPressed(4) = True
    End If
Else
    KeyPressed(4) = False
End If
If (GetKeyState(53) And KEY_PRESSED) Then '5 key was pressed
    If KeyPressed(5) = False Then
    Call HitNotes(5, True)
    KeyPressed(5) = True
    End If
Else
    KeyPressed(5) = False
End If
If (GetKeyState(37) And KEY_PRESSED) Then 'left arrow key was pressed
    If Strum(1) = False Then
    Call StrumNotes
    Strum(1) = True
    End If
Else
    Strum(1) = False
End If
If (GetKeyState(40) And KEY_PRESSED) Then 'down arrow key was pressed
    If Strum(2) = False Then
    Call StrumNotes
    Strum(2) = True
    End If
Else
    Strum(2) = False
End If
If (GetKeyState(39) And KEY_PRESSED) Then 'right arrow key was pressed
    If Strum(3) = False Then
    Call StrumNotes
    Strum(3) = True
    End If
Else
    Strum(3) = False
End If
End Sub

Public Sub UpdateObjects()
Dim i As Integer, k As Integer

For i = NoteMin To NoteCount
    If Note(i).Distance < ScrollY - 100 Then
        If i = NoteMin Then NoteMin = NoteMin + 1
    End If
    If i > NoteMax Then
        If Note(i).Distance <= ScrollY + ScreenHeight Then
            NoteMax = i
        Else
            Exit For
        End If
    End If
Next i
    ScrollY = ScrollY + Notespeed
For i = NoteMin To NoteMax
    'If Note(i).Distance = ScrollY And Note(i).Duration >= 0 And Note(i).Hit = True And Note(i).Deleted = False And KeyPressed(Note(i).XOffset + 1) = True Then

    'Else
    
    If Note(i).Distance < ScrollY - 80 Then
    Note(i).Deleted = True
    If Note(i).Hit = False Then
    Run = 0
    RunBonus = 1
    If Player(ActivePlayer).SoundEffects = True Then frmMain.ESound.SetStreamingVolume SoundInMemory, -800
    End If
    End If
    'End If
    If Note(i).Hit = True Then
        If Note(i).Duration <= 0 Then
            Note(i).Deleted = True
        Else
            If KeyPressed(Note(i).XOffset + 1) = True Then
                Note(i).Distance = ScrollY
                AddScore = AddScore + Notespeed / 12 * RunBonus
                Note(i).Duration = Note(i).Duration - Notespeed
            Else
                Note(i).Deleted = True
            End If
        End If
    End If
Next i
CumScore = Int(Score + AddScore)
End Sub


Public Function HitNotes(Index As Integer, checkhit As Boolean) As Boolean
Dim i As Integer
    For i = NoteMin To NoteMax
        If Note(i).Hit = False Or checkhit = False Then
        If (Note(i).Distance - ScrollY <= (20 + checkhit * 10 + Player(ActivePlayer).ResponseOffset) * BeatsPerMin / 120 And Note(i).Distance - ScrollY >= (-20 - checkhit * 10 + Player(ActivePlayer).ResponseOffset) * BeatsPerMin / 120) And Note(i).XOffset = Index - 1 Then
            HitNotes = True
            If Note(i).Duration > 0 Then Note(i).Duration = Note(i).Distance + Note(i).Duration - ScrollY
            If Note(i).Duration < 0 Then Note(i).Duration = 0
            If Note(i).Hit = False Then
                Note(i).Hit = True
                SpinState(Note(i).XOffset + 1) = 3
                Score = Score + 20 * RunBonus
                Run = Run + 1
                RunBonus = Int(Run / 16) + 1
                If RunBonus > 4 Then RunBonus = 4
            End If
            Exit For
        End If
        End If
    Next i
    If HitNotes = False Then
    Run = 0
    RunBonus = 1
    End If
End Function

Public Sub StrumNotes()
Dim k As Integer
For k = 1 To 5
    If KeyPressed(k) = True Then Call HitNotes(k, False)
Next k
End Sub
