Attribute VB_Name = "Module3"
Public SoundInMemory As Long
'Public PixPerBeat As Integer
Public Sub LoadSong(FileName As String)
Dim temp, i As Integer
Open App.Path & "\Level Files\" & FileName & ".dat" For Input As #1
Line Input #1, temp
Line Input #1, temp
SongFileName = temp
Line Input #1, temp
BeatsPerMin = temp
Line Input #1, temp
NoteCount = temp
ReDim Note(1 To NoteCount)
For i = 1 To NoteCount
Line Input #1, temp
Note(i).Distance = temp * 120
Line Input #1, temp
Note(i).Duration = temp * 120
Line Input #1, temp
Note(i).XOffset = temp
Note(i).Deleted = False
Note(i).Hit = False
Next i
Close #1
End Sub
