Attribute VB_Name = "Module1"
Option Explicit
Public Box() As Box, Boxes As Integer, BoxMin As Integer, BoxDC(1 To 2) As Long, BoxMaskDC(1 To 2) As Long
Public Player As Player, PlayerDC(1 To 2) As Long, MaskDC(1 To 2) As Long
Public Lives As Integer, BoxScore As Integer, Score As Integer
Public BackgroundDC As Long, BackBuffDC As Long
Type Box
    x As Single
    y As Single
    xMove As Single
    yMove As Single
    Deleted As Boolean
    DC As Integer '1 or 2
End Type
Type Player
    x As Single
    y As Single
    xMove As Single
    yMove As Single
    Jump As Boolean
    OnBox As Integer
    Position As Integer ' 1 = right, 2 = left
End Type

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
Const LR_DEFAULTSIZE As Long = &H40
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


Public Sub GenerateBox()
Static lastY As Integer
If lastY = 0 Then lastY = 298
Boxes = Boxes + 1
ReDim Preserve Box(BoxMin To Boxes)
With Box(Boxes)
    .DC = 1
    .Deleted = False
    .y = Int(Rnd * (500 - lastY) + lastY - 100)
    If Box(Boxes).y > 400 Then Box(Boxes).y = 400
    If Box(Boxes).y < 96 Then Box(Boxes).y = 96
If Int(Rnd * 2) = 0 Then
    .xMove = -3
    .x = 600
    .DC = 2
Else
    .xMove = 3
    .x = -130
    .DC = 1
End If
    .yMove = 0
End With
lastY = Box(Boxes).y
End Sub
