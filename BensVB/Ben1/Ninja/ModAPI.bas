Attribute VB_Name = "Module2"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, _
ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'Public Declare Function Sleep Lib "user32" (ByVal Duration) As Long
'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTSIZE As Long = &H40
'**PlaySound Constants**
Const SND_FILENAME = &H20000 'Sound is a filename
Const SND_ASYNC = &H1 'Sound is played asynchronously
Const SND_RESOURCE = &H40004 'Name is a resource id
Const SND_SYNC = &H0 'play synchronously

Public Enum MMApiErrors
    WAVE_NOT_PLAYED = vbObjectError + 700
End Enum

Public Function WinPlaySound(ByVal pSound As String, ByVal flFile As Boolean, ByVal flAsync As Boolean) As Boolean
Dim llngRet As Long
Dim llngMode As Long

If flAsync = True Then
    llngMode = SND_ASYNC
Else
    llngMode = SND_SYNC
End If

If flFile = True Then
    llngRet = PlaySound(pSound, 0, SND_FILENAME Or llngMode)
Else
    llngRet = PlaySound(pSound, App.hInstance, SND_RESOURCE Or llngMode)
End If

If llngRet = 1 Then
    WinPlaySound = True
Else
    Err.Raise MMApiErrors.WAVE_NOT_PLAYED, "WinPlaySound", "Wave sound not played."
    WinPlaySound = False
End If
End Function


'IN: FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

'If DC < 1 Then
'    GenerateDC = 0
'    Exit Function
'End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Public Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function

'Loads a sound from file into memory
Public Function LoadSound(FileName As String) As String
On Error GoTo Error_Handler

Dim FreeFileNumber As Integer
Dim SoundBuffer As String

FreeFileNumber = FreeFile
SoundBuffer = Space$(FileLen(FileName)) 'Make room for sound file

Open FileName For Binary As #FreeFileNumber
         
         Get #FreeFileNumber, , SoundBuffer
         
Close FreeFileNumber

'remove wasted spaces
LoadSound = Trim$(SoundBuffer)
    
Error_Handler:

    Select Case Err
    
        Case 0 'No error
        
        Case Else
            
            MsgBox Err.Description & vbCrLf & "Error Number: " & Err.Number
            Err.Clear
            LoadSound = ""
    End Select
   
End Function


